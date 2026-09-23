import os
import pytest
from app import create_app
from app.models.schema import db, User, Exam, Question, Attempt, Answer, AIFeedback
from app.services.scoring import evaluate_answer, grade_attempt_submission
from app.services.vector_service import index_question, search_similar_questions
from app.services.ai_service import generate_ai_questions, generate_personalized_ai_feedback

from config import Config

class TestConfig(Config):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///:memory:'

@pytest.fixture
def app():
    app = create_app(TestConfig)
    with app.app_context():
        db.create_all()
        yield app
        db.session.remove()
        db.drop_all()

@pytest.fixture
def client(app):
    return app.test_client()

def test_user_creation_and_passwords(app):
    with app.app_context():
        u = User(name="Test User", email="test@example.com", role="student")
        u.set_password("secret123")
        db.session.add(u)
        db.session.commit()

        user = User.query.filter_by(email="test@example.com").first()
        assert user is not None
        assert user.check_password("secret123") is True
        assert user.check_password("wrongpass") is False
        assert user.is_admin() is False

def test_deterministic_scoring(app):
    # Single Choice
    q_single = Question(
        text="What is 2 + 2?",
        question_type="single_choice",
        correct_answers_json='["4"]',
        points=1.0
    )
    is_corr, pts, fb = evaluate_answer(q_single, "4")
    assert is_corr is True
    assert pts == 1.0

    is_corr_wrong, pts_wrong, _ = evaluate_answer(q_single, "5")
    assert is_corr_wrong is False
    assert pts_wrong == 0.0

    # Multi Choice
    q_multi = Question(
        text="Select prime numbers",
        question_type="multi_choice",
        correct_answers_json='["2", "3"]',
        points=2.0
    )
    is_corr_m, pts_m, _ = evaluate_answer(q_multi, ["2", "3"])
    assert is_corr_m is True
    assert pts_m == 2.0

def test_exam_start_and_autosave_api(client, app):
    # Create admin and student
    with app.app_context():
        admin = User(name="Admin", email="admin@test.com", role="admin")
        admin.set_password("pass")
        student = User(name="Student", email="student@test.com", role="student")
        student.set_password("pass")
        db.session.add_all([admin, student])
        db.session.commit()

        exam = Exam(title="Test Exam", duration_minutes=15, is_published=True, created_by_id=admin.id)
        db.session.add(exam)
        db.session.commit()

        q1 = Question(exam_id=exam.id, text="Q1?", question_type="single_choice", points=1.0)
        q1.correct_answers = ["A"]
        db.session.add(q1)
        db.session.commit()
        exam_id = exam.id
        q1_id = q1.id

    # Log in student
    client.post('/login', data={'email': 'student@test.com', 'password': 'pass'})

    # Start exam
    res = client.post(f'/exam/{exam_id}/start', follow_redirects=True)
    assert res.status_code == 200

    with app.app_context():
        attempt = Attempt.query.first()
        assert attempt is not None
        assert attempt.remaining_seconds() > 0
        attempt_id = attempt.id

    # Test background Auto-save API
    res_save = client.post('/api/exam/save-answer', json={
        'attempt_id': attempt_id,
        'question_id': q1_id,
        'response': 'A'
    })
    assert res_save.status_code == 200
    data = res_save.get_json()
    assert data['status'] == 'success'

    # Submit exam
    res_sub = client.post(f'/exam/attempt/{attempt_id}/submit', follow_redirects=True)
    assert res_sub.status_code == 200

    with app.app_context():
        att_graded = Attempt.query.get(attempt_id)
        assert att_graded.status == 'graded'
        assert att_graded.score == 1.0
        assert att_graded.passed is True

def test_head_role_and_user_management(client, app):
    with app.app_context():
        head = User(name="Head Admin", email="head@test.com", role="head")
        head.set_password("headpass")
        target_student = User(name="Target Student", email="target@test.com", role="student")
        target_student.set_password("pass")
        db.session.add(head)
        db.session.add(target_student)
        db.session.commit()
        head_id = head.id
        target_id = target_student.id

    # 1. Login as Head
    res_login = client.post('/login', data={'email': 'head@test.com', 'password': 'headpass'}, follow_redirects=True)
    assert res_login.status_code == 200
    assert b"Executive Command Center" in res_login.data

    # 2. View User Management
    res_users = client.get('/head/users')
    assert res_users.status_code == 200
    assert b"Target Student" in res_users.data

    # 3. Promote Student to Admin
    res_role = client.post(f'/head/users/{target_id}/role', data={'role': 'admin'}, follow_redirects=True)
    assert res_role.status_code == 200
    with app.app_context():
        promoted = User.query.get(target_id)
        assert promoted.role == 'admin'

    # 4. Prevent deleting self
    res_self_del = client.post(f'/head/users/{head_id}/delete', follow_redirects=True)
    assert res_self_del.status_code == 200
    with app.app_context():
        assert User.query.get(head_id) is not None

    # 5. Delete other user
    res_del = client.post(f'/head/users/{target_id}/delete', follow_redirects=True)
    assert res_del.status_code == 200
    with app.app_context():
        assert User.query.get(target_id) is None


def test_proctoring_integrity_events_and_certificate(client, app):
    with app.app_context():
        creator = User(name="Faculty", email="faculty@test.com", role="admin")
        creator.set_password("pass")
        student = User(name="Candidate", email="candidate@test.com", role="student")
        student.set_password("candpass")
        db.session.add(creator)
        db.session.add(student)
        db.session.commit()

        exam = Exam(
            title="Integrity Tested Exam",
            description="Test exam",
            duration_minutes=15,
            pass_mark=50.0,
            is_published=True,
            created_by_id=creator.id
        )
        db.session.add(exam)
        db.session.commit()

        q = Question(
            exam_id=exam.id,
            text="Is Python dynamic?",
            question_type="true_false",
            options_json='["True", "False"]',
            correct_answers_json='["True"]',
            points=10.0
        )
        db.session.add(q)
        db.session.commit()
        exam_id = exam.id
        q_id = q.id

    # Login as student
    client.post('/login', data={'email': 'candidate@test.com', 'password': 'candpass'}, follow_redirects=True)

    # Start exam
    client.post(f'/exam/{exam_id}/start', follow_redirects=True)
    with app.app_context():
        cand = User.query.filter_by(email="candidate@test.com").first()
        att = Attempt.query.filter_by(user_id=cand.id).first()
        att_id = att.id
        assert att.integrity_score == 100

    # Log tab switch integrity infraction
    res_event = client.post('/api/exam/integrity-event', json={
        'attempt_id': att_id,
        'event_type': 'tab_switch',
        'details': 'Browser tab changed or minimized'
    })
    assert res_event.status_code == 200
    data = res_event.get_json()
    assert data['integrity_score'] == 95

    # Answer and submit
    client.post('/api/exam/save-answer', json={
        'attempt_id': att_id,
        'question_id': q_id,
        'response': 'True'
    })
    client.post(f'/exam/attempt/{att_id}/submit', follow_redirects=True)

    # View certificate
    res_cert = client.get(f'/exam/attempt/{att_id}/certificate')
    assert res_cert.status_code == 200
    assert b"Certificate of Achievement" in res_cert.data
    assert b"OES-CERT-2026-" in res_cert.data



