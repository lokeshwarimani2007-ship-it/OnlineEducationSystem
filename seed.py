from app import create_app
from app.models.schema import db, User, Exam, Question, Attempt, Answer
from app.services.vector_service import index_all_questions

def seed_database():
    app = create_app()
    with app.app_context():
        # Clear existing tables
        db.drop_all()
        db.create_all()

        print("Seeding database...")

        # 0. Head of Department / Examination
        head = User(
            name="Dr. Evelyn Vance (Head of Examinations)",
            email="head@oes.com",
            role="head"
        )
        head.set_password("head123")
        db.session.add(head)

        # 1. Admin User
        admin = User(
            name="System Administrator",
            email="admin@oes.com",
            role="admin"
        )
        admin.set_password("admin123")
        db.session.add(admin)

        # 2. Student Users
        student1 = User(
            name="Alice Johnson",
            email="student@oes.com",
            role="student"
        )
        student1.set_password("student123")
        
        student2 = User(
            name="Bob Smith",
            email="bob@oes.com",
            role="student"
        )
        student2.set_password("student123")

        legacy_admin = User(name="Master Admin", email="master@gmail.com", role="admin")
        legacy_admin.set_password("admin123")

        legacy_student = User(name="Student Alpha", email="student@gmail.com", role="student")
        legacy_student.set_password("student123")

        db.session.add(student1)
        db.session.add(student2)
        db.session.add(legacy_admin)
        db.session.add(legacy_student)
        db.session.commit()

        # 3. Create Sample Exams
        exam1 = Exam(
            title="Python & Web Development Certification",
            description="Comprehensive assessment covering core Python 3 concepts, Flask web architecture, and HTTP standards.",
            duration_minutes=30,
            pass_mark=60.0,
            is_published=True,
            created_by_id=admin.id
        )
        
        exam2 = Exam(
            title="Database Systems & Vector Databases",
            description="Test your understanding of relational databases (SQL, ORM) and vector embeddings with ChromaDB.",
            duration_minutes=20,
            pass_mark=50.0,
            is_published=True,
            created_by_id=admin.id
        )

        db.session.add(exam1)
        db.session.add(exam2)
        db.session.commit()

        # 4. Questions for Exam 1
        q1 = Question(
            exam_id=exam1.id,
            text="Which of the following Python data structures is mutable?",
            question_type="single_choice",
            topic="Python Fundamentals",
            difficulty="easy",
            points=1.0,
            explanation="Lists are mutable in Python, whereas tuples, strings, and frozensets are immutable."
        )
        q1.options = ["Tuple", "List", "String", "Frozenset"]
        q1.correct_answers = ["List"]

        q2 = Question(
            exam_id=exam1.id,
            text="Select all valid HTTP Request Methods supported by standard web applications:",
            question_type="multi_choice",
            topic="Web Architecture",
            difficulty="medium",
            points=2.0,
            explanation="GET, POST, PUT, and DELETE are standard HTTP methods."
        )
        q2.options = ["GET", "POST", "FETCH", "DELETE"]
        q2.correct_answers = ["GET", "POST", "DELETE"]

        q3 = Question(
            exam_id=exam1.id,
            text="In Flask web architecture, route decorators execute synchronous functions on the main loop by default.",
            question_type="true_false",
            topic="Flask Framework",
            difficulty="easy",
            points=1.0,
            explanation="Flask WSGI views handle incoming requests synchronously per worker process."
        )
        q3.options = ["True", "False"]
        q3.correct_answers = ["True"]

        q4 = Question(
            exam_id=exam1.id,
            text="What built-in Python function is used to return the length of an iterable object?",
            question_type="short_answer",
            topic="Python Fundamentals",
            difficulty="easy",
            points=1.0,
            explanation="The len() function returns the total number of items in an object."
        )
        q4.correct_answers = ["len", "len()"]

        # Questions for Exam 2
        q5 = Question(
            exam_id=exam2.id,
            text="What type of search does ChromaDB utilize to retrieve contextually similar documents?",
            question_type="single_choice",
            topic="Vector Databases",
            difficulty="medium",
            points=1.0,
            explanation="ChromaDB uses vector similarity search (cosine distance or Euclidean distance) on high-dimensional embeddings."
        )
        q5.options = ["Keyword Exact Search", "Vector Similarity Search", "Regex Search", "B-Tree Indexing"]
        q5.correct_answers = ["Vector Similarity Search"]

        q6 = Question(
            exam_id=exam2.id,
            text="Which SQL clause is used to filter records before grouping occurs?",
            question_type="single_choice",
            topic="SQL & Databases",
            difficulty="medium",
            points=1.0,
            explanation="The WHERE clause filters rows before GROUP BY aggregation, while HAVING filters aggregated groups."
        )
        q6.options = ["HAVING", "WHERE", "ORDER BY", "GROUP BY"]
        q6.correct_answers = ["WHERE"]

        db.session.add_all([q1, q2, q3, q4, q5, q6])
        db.session.commit()

        # 5. Index questions in ChromaDB vector store
        all_qs = Question.query.all()
        index_all_questions(all_qs)

        print(f"Successfully seeded {User.query.count()} users, {Exam.query.count()} exams, and {Question.query.count()} questions into SQLite and ChromaDB!")

if __name__ == '__main__':
    seed_database()
