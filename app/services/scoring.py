from app.models.schema import db, Answer

def evaluate_answer(question, student_response):
    """
    Evaluates student response against question's correct answers.
    Returns tuple: (is_correct: bool, points_awarded: float, feedback: str)
    """
    if student_response is None or student_response == "":
        return False, 0.0, "No response provided."

    q_type = question.question_type
    correct = question.correct_answers
    max_pts = float(question.points)

    if q_type in ['single_choice', 'true_false']:
        # Compare as single string comparison
        resp_str = str(student_response).strip().lower()
        correct_str = str(correct[0] if correct else "").strip().lower()
        if resp_str == correct_str:
            return True, max_pts, "Correct answer!"
        return False, 0.0, f"Incorrect. Correct answer: {correct[0] if correct else 'N/A'}"

    elif q_type == 'multi_choice':
        # Student response expected as list or string
        if isinstance(student_response, str):
            student_list = [s.strip().lower() for s in student_response.split(',')]
        elif isinstance(student_response, list):
            student_list = [str(s).strip().lower() for s in student_response]
        else:
            student_list = [str(student_response).strip().lower()]

        correct_list = [str(c).strip().lower() for c in correct]

        # Exact set match check
        if set(student_list) == set(correct_list):
            return True, max_pts, "All correct choices selected!"
        
        # Partial credit rule: correct selections - incorrect selections
        correct_hits = sum(1 for item in student_list if item in correct_list)
        wrong_hits = sum(1 for item in student_list if item not in correct_list)
        if len(correct_list) > 0:
            fraction = max(0.0, (correct_hits - wrong_hits) / len(correct_list))
            earned = round(max_pts * fraction, 2)
            if earned > 0:
                return False, earned, f"Partially correct ({earned}/{max_pts} pts). Correct choices: {', '.join(correct)}"
        return False, 0.0, f"Incorrect choices. Correct choices: {', '.join(correct)}"

    elif q_type == 'short_answer':
        resp_str = str(student_response).strip().lower()
        correct_items = [str(c).strip().lower() for c in correct]
        
        # Exact or substring match in acceptable list of key answers
        if any(resp_str == item or item in resp_str for item in correct_items if item):
            return True, max_pts, "Short answer matched expected key!"
        
        return False, 0.0, f"Answer did not match key phrases: {', '.join(correct)}"

    return False, 0.0, "Unsupported question type."


def grade_attempt_submission(attempt):
    """
    Grades an entire attempt server-side and calculates total score.
    """
    exam = attempt.exam
    questions = exam.questions
    
    # Build map of saved answers
    existing_answers = {ans.question_id: ans for ans in attempt.answers}
    
    total_earned = 0.0
    total_possible = sum(q.points for q in questions)
    
    for q in questions:
        ans_obj = existing_answers.get(q.id)
        if not ans_obj:
            ans_obj = Answer(
                attempt_id=attempt.id,
                question_id=q.id,
                student_response_json=None,
                is_correct=False,
                points_awarded=0.0,
                ai_feedback="No response submitted."
            )
            db.session.add(ans_obj)
        else:
            resp = ans_obj.student_response
            is_corr, pts, fb = evaluate_answer(q, resp)
            ans_obj.is_correct = is_corr
            ans_obj.points_awarded = pts
            ans_obj.ai_feedback = fb
            total_earned += pts

    attempt.score = round(total_earned, 2)
    attempt.max_score = round(total_possible, 2)
    
    # Calculate pass threshold
    percentage = (total_earned / total_possible * 100.0) if total_possible > 0 else 0.0
    attempt.passed = percentage >= exam.pass_mark
    attempt.status = 'graded'
    db.session.commit()
    
    return attempt
