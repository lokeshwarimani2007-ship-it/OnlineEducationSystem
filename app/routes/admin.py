from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.routes.auth import admin_required
from app.models.schema import db, User, Exam, Question, Attempt, Answer, AIFeedback
from app.services.vector_service import index_question, index_all_questions, search_similar_questions
from app.services.ai_service import generate_ai_questions

admin_bp = Blueprint('admin', __name__)

@admin_bp.route('/dashboard')
@login_required
@admin_required
def dashboard():
    total_exams = Exam.query.count()
    published_exams = Exam.query.filter_by(is_published=True).count()
    total_questions = Question.query.count()
    total_students = User.query.filter_by(role='student').count()
    total_attempts = Attempt.query.filter(Attempt.status != 'in_progress').count()
    
    attempts = Attempt.query.filter(Attempt.status != 'in_progress').all()
    avg_score = round(sum(a.score for a in attempts) / len(attempts), 1) if attempts else 0.0
    pass_rate = round(sum(1 for a in attempts if a.passed) / len(attempts) * 100, 1) if attempts else 0.0

    recent_attempts = Attempt.query.order_by(Attempt.start_time.desc()).limit(10).all()

    return render_template(
        'admin/dashboard.html',
        total_exams=total_exams,
        published_exams=published_exams,
        total_questions=total_questions,
        total_students=total_students,
        total_attempts=total_attempts,
        avg_score=avg_score,
        pass_rate=pass_rate,
        recent_attempts=recent_attempts
    )


# --- QUESTION BANK MANAGEMENT ---

@admin_bp.route('/questions')
@login_required
@admin_required
def questions():
    topic_filter = request.args.get('topic', '').strip()
    difficulty_filter = request.args.get('difficulty', '').strip()
    search_query = request.args.get('q', '').strip()

    query = Question.query
    if topic_filter:
        query = query.filter_by(topic=topic_filter)
    if difficulty_filter:
        query = query.filter_by(difficulty=difficulty_filter)
    if search_query:
        query = query.filter(Question.text.ilike(f'%{search_query}%'))

    questions_list = query.order_by(Question.created_at.desc()).all()
    
    # Get distinct topics
    topics = db.session.query(Question.topic).distinct().all()
    topics_list = [t[0] for t in topics if t[0]]

    return render_template(
        'admin/questions.html',
        questions=questions_list,
        topics=topics_list,
        selected_topic=topic_filter,
        selected_difficulty=difficulty_filter,
        search_query=search_query
    )


@admin_bp.route('/questions/add', methods=['GET', 'POST'])
@login_required
@admin_required
def add_question():
    if request.method == 'POST':
        text = request.form.get('text', '').strip()
        question_type = request.form.get('question_type', 'single_choice')
        topic = request.form.get('topic', 'General').strip()
        difficulty = request.form.get('difficulty', 'medium')
        points = float(request.form.get('points', 1.0))
        explanation = request.form.get('explanation', '').strip()
        exam_id = request.form.get('exam_id')

        options_raw = request.form.getlist('options')
        correct_raw = request.form.getlist('correct_answers')

        # Clean lists
        options = [opt.strip() for opt in options_raw if opt.strip()]
        correct_answers = [ans.strip() for ans in correct_raw if ans.strip()]

        if not text:
            flash('Question text is required.', 'danger')
            return render_template('admin/question_form.html')

        q = Question(
            exam_id=int(exam_id) if exam_id else None,
            text=text,
            question_type=question_type,
            topic=topic,
            difficulty=difficulty,
            points=points,
            explanation=explanation
        )
        q.options = options
        q.correct_answers = correct_answers

        db.session.add(q)
        db.session.commit()

        # Index in ChromaDB
        index_question(q)

        flash('Question created and indexed in vector database successfully!', 'success')
        return redirect(url_for('admin.questions'))

    exams = Exam.query.all()
    return render_template('admin/question_form.html', exams=exams)


@admin_bp.route('/questions/<int:q_id>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_question(q_id):
    question = Question.query.get_or_404(q_id)

    if request.method == 'POST':
        question.text = request.form.get('text', '').strip()
        question.question_type = request.form.get('question_type', 'single_choice')
        question.topic = request.form.get('topic', 'General').strip()
        question.difficulty = request.form.get('difficulty', 'medium')
        question.points = float(request.form.get('points', 1.0))
        question.explanation = request.form.get('explanation', '').strip()
        
        exam_id = request.form.get('exam_id')
        question.exam_id = int(exam_id) if exam_id else None

        options_raw = request.form.getlist('options')
        correct_raw = request.form.getlist('correct_answers')
        
        question.options = [opt.strip() for opt in options_raw if opt.strip()]
        question.correct_answers = [ans.strip() for ans in correct_raw if ans.strip()]

        db.session.commit()

        # Update ChromaDB index
        index_question(question)

        flash('Question updated successfully!', 'success')
        return redirect(url_for('admin.questions'))

    exams = Exam.query.all()
    return render_template('admin/question_form.html', question=question, exams=exams)


@admin_bp.route('/questions/<int:q_id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_question(q_id):
    question = Question.query.get_or_404(q_id)
    db.session.delete(question)
    db.session.commit()
    flash('Question deleted.', 'info')
    return redirect(url_for('admin.questions'))


@admin_bp.route('/questions/ai-generate', methods=['POST'])
@login_required
@admin_required
def ai_generate_questions():
    topic = request.form.get('topic', 'General').strip()
    difficulty = request.form.get('difficulty', 'medium')
    question_type = request.form.get('question_type', 'mixed')
    focus_area = request.form.get('focus_area', '').strip()
    count = int(request.form.get('count', 3))
    exam_id = request.form.get('exam_id')

    generated = generate_ai_questions(
        topic=topic,
        difficulty=difficulty,
        question_type=question_type,
        count=count,
        focus_area=focus_area
    )
    added_count = 0


    for q_data in generated:
        q = Question(
            exam_id=int(exam_id) if exam_id else None,
            text=q_data.get('text'),
            question_type=q_data.get('question_type', 'single_choice'),
            topic=q_data.get('topic', topic),
            difficulty=q_data.get('difficulty', difficulty),
            points=float(q_data.get('points', 1.0)),
            explanation=q_data.get('explanation', '')
        )
        q.options = q_data.get('options', [])
        q.correct_answers = q_data.get('correct_answers', [])

        db.session.add(q)
        db.session.commit()
        index_question(q)
        added_count += 1

    flash(f'Successfully generated and saved {added_count} AI questions for topic "{topic}"!', 'success')
    return redirect(url_for('admin.questions'))


# --- EXAM BUILDER ---

@admin_bp.route('/exams')
@login_required
@admin_required
def exams():
    exams_list = Exam.query.order_by(Exam.created_at.desc()).all()
    return render_template('admin/exams.html', exams=exams_list)


@admin_bp.route('/exams/create', methods=['GET', 'POST'])
@login_required
@admin_required
def create_exam():
    if request.method == 'POST':
        title = request.form.get('title', '').strip()
        description = request.form.get('description', '').strip()
        duration_minutes = int(request.form.get('duration_minutes', 30))
        pass_mark = float(request.form.get('pass_mark', 50.0))
        is_published = 'is_published' in request.form

        if not title:
            flash('Exam title is required.', 'danger')
            return render_template('admin/exam_form.html')

        exam = Exam(
            title=title,
            description=description,
            duration_minutes=duration_minutes,
            pass_mark=pass_mark,
            is_published=is_published,
            created_by_id=current_user.id
        )
        db.session.add(exam)
        db.session.commit()

        flash('Exam created! Now add questions to your exam.', 'success')
        return redirect(url_for('admin.exam_builder', exam_id=exam.id))

    return render_template('admin/exam_form.html')


@admin_bp.route('/exams/<int:exam_id>/builder', methods=['GET', 'POST'])
@login_required
@admin_required
def exam_builder(exam_id):
    exam = Exam.query.get_or_404(exam_id)

    if request.method == 'POST':
        # Update assigned question list
        selected_question_ids = request.form.getlist('question_ids')
        selected_ids = set(int(qid) for qid in selected_question_ids)

        # Clear existing exam questions assignment
        for q in exam.questions:
            q.exam_id = None
        db.session.commit()

        # Assign selected questions
        if selected_ids:
            questions_to_assign = Question.query.filter(Question.id.in_(selected_ids)).all()
            for q in questions_to_assign:
                q.exam_id = exam.id
            db.session.commit()

        flash(f'Updated exam questions ({len(selected_ids)} assigned). Total Points: {exam.total_points()}', 'success')
        return redirect(url_for('admin.exam_builder', exam_id=exam.id))

    assigned_questions = exam.questions
    assigned_ids = set(q.id for q in assigned_questions)
    
    # All bank questions available
    all_bank_questions = Question.query.all()

    return render_template(
        'admin/exam_builder.html',
        exam=exam,
        assigned_questions=assigned_questions,
        assigned_ids=assigned_ids,
        all_bank_questions=all_bank_questions
    )


@admin_bp.route('/exams/<int:exam_id>/toggle-publish', methods=['POST'])
@login_required
@admin_required
def toggle_publish_exam(exam_id):
    exam = Exam.query.get_or_404(exam_id)
    exam.is_published = not exam.is_published
    db.session.commit()
    status_str = "published" if exam.is_published else "unpublished"
    flash(f'Exam "{exam.title}" is now {status_str}.', 'info')
    return redirect(url_for('admin.exams'))


@admin_bp.route('/exams/<int:exam_id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_exam(exam_id):
    exam = Exam.query.get_or_404(exam_id)
    db.session.delete(exam)
    db.session.commit()
    flash('Exam deleted.', 'info')
    return redirect(url_for('admin.exams'))


# --- ANALYTICS CONTROL CENTER ---

@admin_bp.route('/analytics')
@login_required
@admin_required
def analytics():
    completed_attempts = Attempt.query.filter(Attempt.status != 'in_progress').order_by(Attempt.start_time.desc()).all()
    
    # Item Analysis - Question Error Rates
    all_answers = Answer.query.filter(Answer.is_correct != None).all()
    question_stats = {}
    
    for ans in all_answers:
        q_id = ans.question_id
        if q_id not in question_stats:
            question_stats[q_id] = {
                "question": ans.question,
                "total_attempts": 0,
                "correct_count": 0
            }
        question_stats[q_id]["total_attempts"] += 1
        if ans.is_correct:
            question_stats[q_id]["correct_count"] += 1

    item_analysis = []
    for q_id, stats in question_stats.items():
        if stats["question"]:
            acc = round((stats["correct_count"] / stats["total_attempts"]) * 100, 1)
            item_analysis.append({
                "question_id": q_id,
                "text": stats["question"].text,
                "topic": stats["question"].topic,
                "difficulty": stats["question"].difficulty,
                "total_attempts": stats["total_attempts"],
                "accuracy_rate": acc
            })

    # Sort hardest first (lowest accuracy)
    item_analysis.sort(key=lambda x: x["accuracy_rate"])

    return render_template(
        'admin/analytics.html',
        attempts=completed_attempts,
        item_analysis=item_analysis
    )
