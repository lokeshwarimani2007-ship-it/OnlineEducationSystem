import json
from datetime import datetime, timezone
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='student') # 'student' or 'admin'
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    attempts = db.relationship('Attempt', backref='user', lazy=True, cascade="all, delete-orphan")
    exams_created = db.relationship('Exam', backref='creator', lazy=True)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def is_admin(self):
        return self.role in ['admin', 'head']

    def is_head(self):
        return self.role == 'head'

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'email': self.email,
            'role': self.role,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }


class Exam(db.Model):
    __tablename__ = 'exams'
    
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    duration_minutes = db.Column(db.Integer, nullable=False, default=30)
    pass_mark = db.Column(db.Float, nullable=False, default=50.0) # Percentage or points threshold
    is_published = db.Column(db.Boolean, default=False)
    created_by_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    questions = db.relationship('Question', backref='exam', lazy=True, cascade="all, delete-orphan")
    attempts = db.relationship('Attempt', backref='exam', lazy=True, cascade="all, delete-orphan")

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def total_points(self):
        return sum(q.points for q in self.questions)

    def to_dict(self):
        return {
            'id': self.id,
            'title': self.title,
            'description': self.description,
            'duration_minutes': self.duration_minutes,
            'pass_mark': self.pass_mark,
            'is_published': self.is_published,
            'total_questions': len(self.questions),
            'total_points': self.total_points(),
            'created_at': self.created_at.isoformat() if self.created_at else None
        }


class Question(db.Model):
    __tablename__ = 'questions'
    
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, db.ForeignKey('exams.id'), nullable=True) # Nullable for question bank reuse
    text = db.Column(db.Text, nullable=False)
    question_type = db.Column(db.String(30), nullable=False, default='single_choice') 
    # Types: 'single_choice', 'multi_choice', 'true_false', 'short_answer'
    options_json = db.Column(db.Text, nullable=True) # JSON list of options e.g. ["Opt A", "Opt B"]
    correct_answers_json = db.Column(db.Text, nullable=False) # JSON list e.g. ["Opt A"] or ["True"]
    points = db.Column(db.Float, nullable=False, default=1.0)
    explanation = db.Column(db.Text, nullable=True)
    topic = db.Column(db.String(100), nullable=False, default='General')
    difficulty = db.Column(db.String(20), nullable=False, default='medium') # 'easy', 'medium', 'hard'
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    @property
    def options(self):
        if self.options_json:
            try:
                return json.loads(self.options_json)
            except Exception:
                return []
        return []

    @options.setter
    def options(self, val):
        self.options_json = json.dumps(val)

    @property
    def correct_answers(self):
        if self.correct_answers_json:
            try:
                return json.loads(self.correct_answers_json)
            except Exception:
                return []
        return []

    @correct_answers.setter
    def correct_answers(self, val):
        self.correct_answers_json = json.dumps(val)

    def to_dict(self, include_correct=False):
        data = {
            'id': self.id,
            'exam_id': self.exam_id,
            'text': self.text,
            'question_type': self.question_type,
            'options': self.options,
            'points': self.points,
            'topic': self.topic,
            'difficulty': self.difficulty,
            'explanation': self.explanation if include_correct else None
        }
        if include_correct:
            data['correct_answers'] = self.correct_answers
        return data


class Attempt(db.Model):
    __tablename__ = 'attempts'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    exam_id = db.Column(db.Integer, db.ForeignKey('exams.id'), nullable=False)
    start_time = db.Column(db.DateTime, nullable=False, default=lambda: datetime.now(timezone.utc))
    end_time = db.Column(db.DateTime, nullable=False) # Server authoritative target finish time
    submitted_at = db.Column(db.DateTime, nullable=True)
    status = db.Column(db.String(30), nullable=False, default='in_progress') # 'in_progress', 'submitted', 'graded'
    score = db.Column(db.Float, nullable=True, default=0.0)
    max_score = db.Column(db.Float, nullable=False, default=0.0)
    passed = db.Column(db.Boolean, nullable=True, default=False)
    integrity_score = db.Column(db.Integer, default=100)
    certificate_id = db.Column(db.String(64), nullable=True, unique=True)

    answers = db.relationship('Answer', backref='attempt', lazy=True, cascade="all, delete-orphan")
    ai_feedback = db.relationship('AIFeedback', backref='attempt', uselist=False, cascade="all, delete-orphan")
    integrity_events = db.relationship('IntegrityEvent', backref='attempt', lazy=True, cascade="all, delete-orphan")

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def remaining_seconds(self):
        if self.status != 'in_progress':
            return 0
        now = datetime.now(timezone.utc)
        # Handle naive datetime comparisons if start_time/end_time stored naive
        end = self.end_time
        if end.tzinfo is None:
            end = end.replace(tzinfo=timezone.utc)
        rem = (end - now).total_seconds()
        return max(0, int(rem))

    def to_dict(self):
        return {
            'id': self.id,
            'user_id': self.user_id,
            'exam_id': self.exam_id,
            'exam_title': self.exam.title if self.exam else '',
            'start_time': self.start_time.isoformat() if self.start_time else None,
            'end_time': self.end_time.isoformat() if self.end_time else None,
            'submitted_at': self.submitted_at.isoformat() if self.submitted_at else None,
            'status': self.status,
            'score': self.score,
            'max_score': self.max_score,
            'passed': self.passed,
            'integrity_score': self.integrity_score if self.integrity_score is not None else 100,
            'certificate_id': self.certificate_id,
            'remaining_seconds': self.remaining_seconds()
        }


class Answer(db.Model):
    __tablename__ = 'answers'
    
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey('attempts.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('questions.id'), nullable=False)
    student_response_json = db.Column(db.Text, nullable=True) # JSON list or string
    is_correct = db.Column(db.Boolean, nullable=True)
    points_awarded = db.Column(db.Float, default=0.0)
    ai_feedback = db.Column(db.Text, nullable=True)
    saved_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    question = db.relationship('Question', lazy=True)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    @property
    def student_response(self):
        if self.student_response_json:
            try:
                return json.loads(self.student_response_json)
            except Exception:
                return self.student_response_json
        return None

    @student_response.setter
    def student_response(self, val):
        self.student_response_json = json.dumps(val)


class AIFeedback(db.Model):
    __tablename__ = 'ai_feedback'
    
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey('attempts.id'), nullable=False, unique=True)
    summary = db.Column(db.Text, nullable=False)
    strengths_json = db.Column(db.Text, nullable=True)
    improvements_json = db.Column(db.Text, nullable=True)
    recommended_topics_json = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    @property
    def strengths(self):
        return json.loads(self.strengths_json) if self.strengths_json else []

    @strengths.setter
    def strengths(self, val):
        self.strengths_json = json.dumps(val)

    @property
    def improvements(self):
        return json.loads(self.improvements_json) if self.improvements_json else []

    @improvements.setter
    def improvements(self, val):
        self.improvements_json = json.dumps(val)

    @property
    def recommended_topics(self):
        return json.loads(self.recommended_topics_json) if self.recommended_topics_json else []

    @recommended_topics.setter
    def recommended_topics(self, val):
        self.recommended_topics_json = json.dumps(val)


class AICache(db.Model):
    __tablename__ = 'ai_cache'
    
    id = db.Column(db.Integer, primary_key=True)
    prompt_hash = db.Column(db.String(64), unique=True, nullable=False, index=True)
    response_json = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def __init__(self, **kwargs):
        super().__init__(**kwargs)


class IntegrityEvent(db.Model):
    __tablename__ = 'integrity_events'
    
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey('attempts.id'), nullable=False)
    event_type = db.Column(db.String(50), nullable=False) # 'tab_switch', 'window_blur', 'copy_attempt', 'fullscreen_exit'
    details = db.Column(db.Text, nullable=True)
    timestamp = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def to_dict(self):
        return {
            'id': self.id,
            'attempt_id': self.attempt_id,
            'event_type': self.event_type,
            'details': self.details,
            'timestamp': self.timestamp.isoformat() if self.timestamp else None
        }

