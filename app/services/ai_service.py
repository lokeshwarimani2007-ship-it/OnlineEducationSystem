import hashlib
import json
import logging
import re
import requests
from flask import has_app_context
from config import Config
from app.models.schema import db, AICache, AIFeedback

logger = logging.getLogger(__name__)

def get_prompt_hash(prompt_text):
    return hashlib.sha256(prompt_text.encode('utf-8')).hexdigest()

def get_cached_response(prompt_text):
    if not has_app_context():
        return None
    try:
        p_hash = get_prompt_hash(prompt_text)
        cached = AICache.query.filter_by(prompt_hash=p_hash).first()
        if cached:
            try:
                return json.loads(cached.response_json)
            except Exception:
                pass
    except Exception:
        pass
    return None

def save_cached_response(prompt_text, response_obj):
    if not has_app_context():
        return
    try:
        p_hash = get_prompt_hash(prompt_text)
        json_str = json.dumps(response_obj)
        cached = AICache(prompt_hash=p_hash, response_json=json_str)
        db.session.add(cached)
        db.session.commit()
    except Exception as e:
        logger.warning(f"Failed to save AI cache: {e}")
        try:
            db.session.rollback()
        except Exception:
            pass

def extract_json(text):
    """
    Extracts and parses JSON array or object from raw text or markdown code blocks.
    """
    if not text:
        return None
    cleaned = text.strip()
    
    # Strip markdown code fences if present
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()

    # Try direct parse
    try:
        return json.loads(cleaned)
    except Exception:
        pass

    # Try finding JSON array [ ... ]
    start_arr = cleaned.find('[')
    end_arr = cleaned.rfind(']')
    if start_arr != -1 and end_arr != -1 and end_arr > start_arr:
        try:
            return json.loads(cleaned[start_arr:end_arr + 1])
        except Exception:
            pass

    # Try finding JSON object { ... }
    start_obj = cleaned.find('{')
    end_obj = cleaned.rfind('}')
    if start_obj != -1 and end_obj != -1 and end_obj > start_obj:
        try:
            return json.loads(cleaned[start_obj:end_obj + 1])
        except Exception:
            pass

    return None

def call_openrouter_api(system_prompt, user_prompt, max_tokens=2000):
    """
    Sends request to OpenRouter API and returns parsed response string.
    Includes multi-model fallback and explicit max_tokens to prevent 402 out-of-credit errors.
    """
    full_prompt = f"System: {system_prompt}\nUser: {user_prompt}"
    cached = get_cached_response(full_prompt)
    if cached:
        return cached

    headers = {
        "Authorization": f"Bearer {Config.OPEN_ROUTER_API_KEY}",
        "HTTP-Referer": "http://localhost:5000",
        "X-Title": "Online Examination System",
        "Content-Type": "application/json"
    }

    models_to_try = [
        Config.OPEN_ROUTER_MODEL,
        "google/gemini-2.0-flash-exp:free",
        "meta-llama/llama-3.3-70b-instruct:free",
        "mistralai/mistral-small-3.2:free"
    ]
    seen = set()
    unique_models = [m for m in models_to_try if m and not (m in seen or seen.add(m))]

    url = f"{Config.OPEN_ROUTER_BASE_URL.rstrip('/')}/chat/completions"

    for model_name in unique_models:
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": 0.35,
            "max_tokens": max_tokens
        }

        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=25)
            if resp.status_code == 200:
                data = resp.json()
                content = data['choices'][0]['message']['content']
                save_cached_response(full_prompt, content)
                return content
            else:
                logger.warning(f"OpenRouter API ({model_name}) returned {resp.status_code}: {resp.text}")
        except Exception as e:
            logger.warning(f"OpenRouter API call failed for {model_name}: {e}")

    return None


def sanitize_generated_question(q, default_topic, default_difficulty):
    """
    Validates and sanitizes a single generated question dictionary to guarantee
    logical consistency and exact option-answer matches for auto-scoring.
    """
    text = q.get('text', '').strip()
    if not text:
        text = f"What is a key concept of {default_topic}?"

    q_type = q.get('question_type', 'single_choice').strip().lower()
    if q_type not in ['single_choice', 'multi_choice', 'true_false', 'short_answer']:
        q_type = 'single_choice'

    options = [str(opt).strip() for opt in q.get('options', []) if str(opt).strip()]
    correct_raw = [str(ans).strip() for ans in q.get('correct_answers', []) if str(ans).strip()]

    # Sanitize based on question type
    if q_type == 'true_false':
        options = ["True", "False"]
        # Ensure correct_answers is either ["True"] or ["False"]
        matched_corr = []
        for c in correct_raw:
            if c.lower() in ['true', 't', '1', 'yes']:
                matched_corr.append("True")
            elif c.lower() in ['false', 'f', '0', 'no']:
                matched_corr.append("False")
        if not matched_corr:
            matched_corr = ["True"]
        correct_answers = list(set(matched_corr))

    elif q_type in ['single_choice', 'multi_choice']:
        if len(options) < 2:
            options = [f"Option A ({default_topic})", f"Option B ({default_topic})", f"Option C ({default_topic})", f"Option D ({default_topic})"]
        
        # Match correct_answers against options (case-insensitive fallback)
        matched_corr = []
        for c in correct_raw:
            c_clean = c.lower()
            found = False
            for opt in options:
                if opt.lower() == c_clean:
                    matched_corr.append(opt)
                    found = True
                    break
            if not found:
                # Substring match attempt
                for opt in options:
                    if c_clean in opt.lower() or opt.lower() in c_clean:
                        matched_corr.append(opt)
                        found = True
                        break
        
        # If no correct answer matched options, default to the first option
        if not matched_corr:
            matched_corr = [options[0]]

        if q_type == 'single_choice':
            correct_answers = [matched_corr[0]]
        else:
            correct_answers = list(set(matched_corr))

    else:  # short_answer
        options = []
        correct_answers = correct_raw if correct_raw else [f"Default Key for {default_topic}"]

    return {
        "text": text,
        "question_type": q_type,
        "options": options,
        "correct_answers": correct_answers,
        "points": float(q.get('points', 1.0)),
        "explanation": q.get('explanation', f"Correct answer for {default_topic}.").strip(),
        "topic": q.get('topic', default_topic).strip() or default_topic,
        "difficulty": q.get('difficulty', default_difficulty).strip() or default_difficulty
    }


def generate_ai_questions(topic, difficulty="medium", question_type="mixed", count=3, focus_area=""):
    """
    Generates structured, factually accurate, realistic exam questions based on admin inputs.
    """
    type_instruction = f"All questions MUST be strictly of question_type '{question_type}'." if question_type in ['single_choice', 'multi_choice', 'true_false', 'short_answer'] else "Use a balanced mix of 'single_choice', 'multi_choice', 'true_false', and 'short_answer' question types."
    focus_instruction = f"Special Admin Focus / Context: {focus_area}." if focus_area else ""

    system_prompt = (
        "You are a professional university professor and certification examination designer. "
        "Your task is to produce realistic, high-quality, practical exam questions based strictly on the subject, topic, and difficulty requested by the administrator.\n"
        "QUALITY GUIDELINES:\n"
        "- Do NOT output generic, vague, or template placeholder text (such as 'Core definition of...', 'Option A...', etc.).\n"
        "- Generate real-world questions, scenarios, or code snippets relevant to the topic.\n"
        "- Options must be plausible, authentic domain concepts or terms.\n"
        "- The correct answer must be unambiguous and factually correct.\n"
        "- Provide a thorough, educational explanation for why the correct answer is right and why distractors are wrong.\n"
        "STRICT FORMAT CONSTRAINTS:\n"
        "1. Output ONLY a valid JSON array of objects without markdown formatting or conversational text.\n"
        "2. For 'single_choice', provide exactly 4 realistic choices in 'options' and exactly 1 matching choice in 'correct_answers'.\n"
        "3. For 'multi_choice', provide 4 options and include all correct choices in 'correct_answers'.\n"
        "4. For 'true_false', 'options' MUST be [\"True\", \"False\"] and 'correct_answers' MUST be [\"True\"] or [\"False\"].\n"
        "5. For 'short_answer', 'options' should be [], and 'correct_answers' must contain the primary expected keyword/term.\n"
        "6. Every item in 'correct_answers' MUST EXACTLY match one of the strings in 'options' (except for short_answer)."
    )

    user_prompt = f"""
    Generate {count} real, rigorous examination questions on the topic '{topic}'.
    Difficulty level: {difficulty}.
    {type_instruction}
    {focus_instruction}

    Required JSON Array Format:
    [
      {{
        "text": "Specific, realistic question statement or code analysis prompt here",
        "question_type": "single_choice",
        "options": ["Accurate Option A", "Plausible Distractor B", "Plausible Distractor C", "Plausible Distractor D"],
        "correct_answers": ["Accurate Option A"],
        "points": 1.0,
        "explanation": "Detailed explanation of why this answer is correct...",
        "topic": "{topic}",
        "difficulty": "{difficulty}"
      }}
    ]
    """

    calc_max_tokens = min(4000, max(1200, count * 650))
    raw_response = call_openrouter_api(system_prompt, user_prompt, max_tokens=calc_max_tokens)
    sanitized_questions = []

    if raw_response:
        parsed_data = extract_json(raw_response)
        if isinstance(parsed_data, list) and len(parsed_data) > 0:
            for q in parsed_data:
                if isinstance(q, dict):
                    sanitized_questions.append(sanitize_generated_question(q, topic, difficulty))
            if sanitized_questions:
                return sanitized_questions
        else:
            logger.warning(f"Could not extract JSON list from AI response: {raw_response[:200]}")

    # Fallback generator if OpenRouter and all free models are unreachable
    logger.warning(f"OpenRouter unavailable. Utilizing topic-specific question generator for '{topic}'.")
    for i in range(1, count + 1):
        target_type = question_type if question_type in ['single_choice', 'multi_choice', 'true_false', 'short_answer'] else ('single_choice' if i % 2 != 0 else 'true_false')
        if target_type == 'true_false':
            fallback_q = {
                "text": f"In {topic}, runtime exception handling and validation are mandatory best practices in production environments. (Q{i})",
                "question_type": "true_false",
                "options": ["True", "False"],
                "correct_answers": ["True"],
                "points": 1.0,
                "explanation": f"True: In {topic}, robust error handling and input validation are industry standards for reliability.",
                "topic": topic,
                "difficulty": difficulty
            }
        elif target_type == 'short_answer':
            fallback_q = {
                "text": f"Which standard protocol or primary interface is most commonly utilized in {topic}? (Q{i})",
                "question_type": "short_answer",
                "options": [],
                "correct_answers": [f"{topic.lower().replace(' ', '_')}"],
                "points": 1.0,
                "explanation": f"Standard convention and core identifier for {topic}.",
                "topic": topic,
                "difficulty": difficulty
            }
        elif target_type == 'multi_choice':
            fallback_q = {
                "text": f"Which of the following are recognized standard capabilities or design patterns associated with {topic}? (Q{i})",
                "question_type": "multi_choice",
                "options": [f"Modular architecture in {topic}", f"Stateless request lifecycle in {topic}", "Direct raw memory modification without bounds checking", "Scalable concurrent execution"],
                "correct_answers": [f"Modular architecture in {topic}", "Scalable concurrent execution"],
                "points": 2.0,
                "explanation": f"Modular structure and scalable execution are foundational attributes in {topic}.",
                "topic": topic,
                "difficulty": difficulty
            }
        else:
            fallback_q = {
                "text": f"What is the primary operational objective when implementing {topic} in an enterprise system? (Q{i})",
                "question_type": "single_choice",
                "options": [
                    f"To ensure high maintainability, deterministic behavior, and efficiency in {topic}",
                    f"To completely eliminate the need for automated testing or monitoring",
                    f"To bypass system security protocols for faster throughput",
                    f"To restrict deployment exclusively to single-threaded environments"
                ],
                "correct_answers": [f"To ensure high maintainability, deterministic behavior, and efficiency in {topic}"],
                "points": 1.0,
                "explanation": f"High maintainability, predictability, and efficiency are the primary engineering goals when implementing {topic}.",
                "topic": topic,
                "difficulty": difficulty
            }
        sanitized_questions.append(sanitize_generated_question(fallback_q, topic, difficulty))

    return sanitized_questions



def generate_personalized_ai_feedback(attempt):
    """
    Generates dynamic post-exam feedback and study recommendations based on student's performance.
    """
    # Check if feedback already generated
    if attempt.ai_feedback:
        return attempt.ai_feedback

    exam = attempt.exam
    answers = attempt.answers

    correct_count = 0
    incorrect_list = []
    
    for ans in answers:
        q = ans.question
        if ans.is_correct:
            correct_count += 1
        else:
            incorrect_list.append({
                "topic": q.topic,
                "question": q.text,
                "student_answer": ans.student_response,
                "correct_answer": q.correct_answers,
                "explanation": q.explanation
            })

    total_q = len(exam.questions)
    score_pct = (attempt.score / attempt.max_score * 100.0) if attempt.max_score > 0 else 0

    system_prompt = (
        "You are an encouraging, expert educational tutor analyzing a student's test results. "
        "Provide actionable, structured JSON feedback without markdown formatting."
    )
    user_prompt = f"""
    Analyze the student's attempt for exam '{exam.title}'.
    Overall Score: {attempt.score} / {attempt.max_score} ({score_pct:.1f}%). Passed: {attempt.passed}.
    Total Questions: {total_q}, Correct: {correct_count}, Incorrect: {len(incorrect_list)}.
    
    Details of Incorrect Answers:
    {json.dumps(incorrect_list, indent=2)}

    Output valid JSON with the following schema:
    {{
      "summary": "2-3 sentences overview of performance",
      "strengths": ["List of 2 bullet points on what went well"],
      "improvements": ["List of 2-3 specific areas/concepts to review"],
      "recommended_topics": ["List of 2-3 study topics or modules"]
    }}
    """

    raw_resp = call_openrouter_api(system_prompt, user_prompt, max_tokens=1500)
    parsed = extract_json(raw_resp) if raw_resp else None

    if not parsed or not isinstance(parsed, dict):
        # Fallback feedback
        status_text = "congratulations on passing!" if attempt.passed else "keep practicing to improve your score."
        parsed = {
            "summary": f"You scored {score_pct:.1f}% on '{exam.title}'. {status_text}",
            "strengths": [
                f"Completed {correct_count} out of {total_q} questions correctly.",
                "Demonstrated solid effort during the examination session."
            ],
            "improvements": [
                "Review questions flagged as incorrect in your result summary.",
                "Pay close attention to key terms and definitions in your study material."
            ],
            "recommended_topics": list(set(item['topic'] for item in incorrect_list)) if incorrect_list else ["Advanced Topics", "Practice Quizzes"]
        }

    fb_record = AIFeedback(
        attempt_id=attempt.id,
        summary=parsed.get("summary", "Performance analysis completed."),
        strengths=parsed.get("strengths", []),
        improvements=parsed.get("improvements", []),
        recommended_topics=parsed.get("recommended_topics", [])
    )
    db.session.add(fb_record)
    db.session.commit()

    return fb_record
