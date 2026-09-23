from datetime import datetime, timedelta, timezone
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.models.schema import db, Exam, Question, Attempt, Answer, AIFeedback
from app.services.scoring import grade_attempt_submission
from app.services.ai_service import generate_personalized_ai_feedback

student_bp = Blueprint('student', __name__)

@student_bp.route('/dashboard')
@login_required
def dashboard():
    if current_user.is_admin():
        return redirect(url_for('admin.dashboard'))

    # Active published exams
    available_exams = Exam.query.filter_by(is_published=True).all()

    # User attempts
    user_attempts = Attempt.query.filter_by(user_id=current_user.id).order_by(Attempt.start_time.desc()).all()
    
    in_progress = [a for a in user_attempts if a.status == 'in_progress']
    completed = [a for a in user_attempts if a.status in ['submitted', 'graded']]

    # Analytics calculation
    total_taken = len(completed)
    total_passed = sum(1 for a in completed if a.passed)
    avg_score = round(sum(a.score for a in completed) / total_taken, 1) if total_taken > 0 else 0.0

    return render_template(
        'student/dashboard.html',
        available_exams=available_exams,
        in_progress_attempts=in_progress,
        completed_attempts=completed,
        total_taken=total_taken,
        total_passed=total_passed,
        avg_score=avg_score
    )


@student_bp.route('/exam/<int:exam_id>/start', methods=['GET', 'POST'])
@login_required
def start_exam(exam_id):
    exam = Exam.query.get_or_404(exam_id)
    if not exam.is_published and not current_user.is_admin():
        flash('This exam is not currently available.', 'warning')
        return redirect(url_for('student.dashboard'))

    # Check if student already has an active in_progress attempt
    existing_attempt = Attempt.query.filter_by(
        user_id=current_user.id,
        exam_id=exam.id,
        status='in_progress'
    ).first()

    if existing_attempt:
        # Check if time expired
        if existing_attempt.remaining_seconds() <= 0:
            grade_attempt_submission(existing_attempt)
            generate_personalized_ai_feedback(existing_attempt)
            flash('Your previous attempt timed out and was automatically submitted.', 'info')
            return redirect(url_for('student.view_result', attempt_id=existing_attempt.id))
        return redirect(url_for('student.take_exam', attempt_id=existing_attempt.id))

    # Create new attempt
    now = datetime.now(timezone.utc)
    end_time = now + timedelta(minutes=exam.duration_minutes)
    
    attempt = Attempt(
        user_id=current_user.id,
        exam_id=exam.id,
        start_time=now,
        end_time=end_time,
        status='in_progress',
        score=0.0,
        max_score=exam.total_points(),
        passed=False
    )
    db.session.add(attempt)
    db.session.commit()

    return redirect(url_for('student.take_exam', attempt_id=attempt.id))


@student_bp.route('/exam/attempt/<int:attempt_id>')
@login_required
def take_exam(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)

    if attempt.user_id != current_user.id and not current_user.is_admin():
        flash('Unauthorized access to attempt.', 'danger')
        return redirect(url_for('student.dashboard'))

    if attempt.status != 'in_progress':
        return redirect(url_for('student.view_result', attempt_id=attempt.id))

    # Server timer check
    remaining_sec = attempt.remaining_seconds()
    if remaining_sec <= 0:
        grade_attempt_submission(attempt)
        generate_personalized_ai_feedback(attempt)
        flash('Time has expired! Your exam was automatically submitted.', 'warning')
        return redirect(url_for('student.view_result', attempt_id=attempt.id))

    exam = attempt.exam
    questions = exam.questions

    # Fetch existing saved answers as a map question_id -> answer
    saved_answers = {ans.question_id: ans.student_response for ans in attempt.answers}

    return render_template(
        'student/exam_interface.html',
        attempt=attempt,
        exam=exam,
        questions=questions,
        saved_answers=saved_answers,
        remaining_seconds=remaining_sec
    )


@student_bp.route('/exam/attempt/<int:attempt_id>/submit', methods=['POST'])
@login_required
def submit_exam(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)

    if attempt.user_id != current_user.id and not current_user.is_admin():
        flash('Unauthorized attempt submission.', 'danger')
        return redirect(url_for('student.dashboard'))

    if attempt.status != 'in_progress':
        return redirect(url_for('student.view_result', attempt_id=attempt.id))

    # Grade attempt server-side
    attempt.submitted_at = datetime.now(timezone.utc)
    grade_attempt_submission(attempt)

    # Trigger AI Feedback generation
    try:
        generate_personalized_ai_feedback(attempt)
    except Exception as e:
        pass

    flash('Exam submitted successfully!', 'success')
    return redirect(url_for('student.view_result', attempt_id=attempt.id))


@student_bp.route('/exam/attempt/<int:attempt_id>/result')
@login_required
def view_result(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)

    if attempt.user_id != current_user.id and not current_user.is_admin():
        flash('Unauthorized viewing of results.', 'danger')
        return redirect(url_for('student.dashboard'))

    if attempt.status == 'in_progress':
        return redirect(url_for('student.take_exam', attempt_id=attempt.id))

    # Ensure AI feedback exists
    ai_fb = generate_personalized_ai_feedback(attempt)

    exam = attempt.exam
    questions = exam.questions
    answers_map = {ans.question_id: ans for ans in attempt.answers}

    score_percentage = (attempt.score / attempt.max_score * 100.0) if attempt.max_score > 0 else 0.0

    return render_template(
        'student/result.html',
        attempt=attempt,
        exam=exam,
        questions=questions,
        answers_map=answers_map,
        ai_feedback=ai_fb,
        score_percentage=score_percentage
    )


@student_bp.route('/exam/attempt/<int:attempt_id>/certificate')
@login_required
def view_certificate(attempt_id):
    import uuid
    attempt = Attempt.query.get_or_404(attempt_id)

    if attempt.user_id != current_user.id and not current_user.is_admin():
        flash('Unauthorized access to certificate.', 'danger')
        return redirect(url_for('student.dashboard'))

    if not attempt.passed:
        flash('Certificates are issued exclusively for passed examinations.', 'warning')
        return redirect(url_for('student.view_result', attempt_id=attempt.id))

    if not attempt.certificate_id:
        attempt.certificate_id = f"OES-CERT-2026-{uuid.uuid4().hex[:10].upper()}"
        db.session.commit()

    score_percentage = (attempt.score / attempt.max_score * 100.0) if attempt.max_score > 0 else 0.0

    return render_template(
        'student/certificate.html',
        attempt=attempt,
        exam=attempt.exam,
        student=attempt.user,
        score_percentage=score_percentage
    )

