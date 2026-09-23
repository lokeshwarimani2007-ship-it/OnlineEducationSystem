from datetime import datetime, timezone
from flask import Blueprint, request, jsonify
from flask_login import login_required, current_user
from app.models.schema import db, Attempt, Answer, Question, IntegrityEvent
from app.services.vector_service import search_similar_questions
from app.services.ai_service import generate_ai_questions

api_bp = Blueprint('api', __name__)

@api_bp.route('/exam/save-answer', methods=['POST'])
@login_required
def save_answer():
    data = request.get_json() or {}
    attempt_id = data.get('attempt_id')
    question_id = data.get('question_id')
    response_val = data.get('response')

    if not attempt_id or not question_id:
        return jsonify({'error': 'Missing attempt_id or question_id'}), 400

    attempt = Attempt.query.get(attempt_id)
    if not attempt or (attempt.user_id != current_user.id and not current_user.is_admin()):
        return jsonify({'error': 'Unauthorized'}), 403

    if attempt.status != 'in_progress':
        return jsonify({'error': 'Attempt is no longer in progress', 'status': attempt.status}), 400

    # Server timer check
    if attempt.remaining_seconds() <= 0:
        return jsonify({'error': 'Exam timer expired', 'expired': True}), 400

    # Find or create Answer record
    ans_obj = Answer.query.filter_by(attempt_id=attempt.id, question_id=question_id).first()
    if not ans_obj:
        ans_obj = Answer(attempt_id=attempt.id, question_id=question_id)
        db.session.add(ans_obj)

    ans_obj.student_response = response_val
    ans_obj.saved_at = datetime.now(timezone.utc)
    db.session.commit()

    return jsonify({
        'status': 'success',
        'saved_at': ans_obj.saved_at.isoformat(),
        'remaining_seconds': attempt.remaining_seconds()
    })


@api_bp.route('/exam/timer/<int:attempt_id>')
@login_required
def get_timer(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)
    if attempt.user_id != current_user.id and not current_user.is_admin():
        return jsonify({'error': 'Unauthorized'}), 403

    return jsonify({
        'status': attempt.status,
        'remaining_seconds': attempt.remaining_seconds()
    })


@api_bp.route('/admin/questions/semantic-search', methods=['POST'])
@login_required
def semantic_search():
    if not current_user.is_admin():
        return jsonify({'error': 'Unauthorized'}), 403

    data = request.get_json() or {}
    query_text = data.get('query', '').strip()
    topic = data.get('topic')

    if not query_text:
        return jsonify({'results': []})

    results = search_similar_questions(query_text=query_text, n_results=5, topic_filter=topic)
    return jsonify({'results': results})


@api_bp.route('/exam/integrity-event', methods=['POST'])
@login_required
def log_integrity_event():
    data = request.get_json() or {}
    attempt_id = data.get('attempt_id')
    event_type = data.get('event_type', 'window_blur')
    details = data.get('details', '')

    if not attempt_id:
        return jsonify({'error': 'Missing attempt_id'}), 400

    attempt = Attempt.query.get(attempt_id)
    if not attempt or (attempt.user_id != current_user.id and not current_user.is_admin()):
        return jsonify({'error': 'Unauthorized'}), 403

    if attempt.status != 'in_progress':
        return jsonify({'error': 'Attempt is not in progress'}), 400

    # Calculate penalty
    penalty = 5
    if event_type == 'copy_attempt':
        penalty = 10
    elif event_type in ['tab_switch', 'window_blur']:
        penalty = 5
    elif event_type == 'fullscreen_exit':
        penalty = 8

    current_score = attempt.integrity_score if attempt.integrity_score is not None else 100
    attempt.integrity_score = max(0, current_score - penalty)

    event = IntegrityEvent(
        attempt_id=attempt.id,
        event_type=event_type,
        details=details
    )
    db.session.add(event)
    db.session.commit()

    return jsonify({
        'status': 'logged',
        'integrity_score': attempt.integrity_score,
        'event_id': event.id
    })
