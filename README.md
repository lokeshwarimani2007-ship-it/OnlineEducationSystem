# Online Examination System (OES)
An Online Education & Examination Platform

A full-stack, web-based Online Examination System (OES) built with Python (Flask), SQLAlchemy, SQLite, ChromaDB vector search, OpenRouter AI integration, and responsive glassmorphism UI.

## Features

- **Auth & Role Management**: Multi-role authentication (Students vs. Administrators/Educators).
- **Secure Exam Engine**:
  - Server-authoritative timer preventing client manipulation.
  - Background periodic response auto-saving.
  - Interactive Question Navigator with Flag for Review system.
  - Deterministic server-side grading for single/multi choice, true/false, and short answer.
- **Admin Control Center**:
  - Question Bank CRUD with topic & difficulty tagging.
  - Exam Builder for creating, timing, and publishing exams.
  - Performance Analytics & Item Analysis (identifying hardest/most missed questions).
- **AI & Vector DB Integration**:
  - Local **ChromaDB** store for semantic question indexing and vector search.
  - **OpenRouter API** integration for automated AI question generation.
  - Personalized post-exam performance summary, strengths, area breakdowns, and study recommendations.
  - Local SQLite response caching to save API tokens and reduce latency.

---

## Directory Structure

```
oes/
├── app/
│   ├── models/        # SQLAlchemy schemas (User, Exam, Question, Attempt, Answer, AIFeedback, AICache)
│   ├── routes/        # Auth, Student, Admin, and API route blueprints
│   ├── services/      # Scoring engine, OpenRouter AI wrapper, ChromaDB vector search
│   ├── static/        # Glassmorphism CSS design system & client-side exam JS
│   └── templates/     # Jinja2 HTML templates for auth, student, admin, and results
├── data/
│   ├── oes.db         # SQLite database file
│   └── chroma/        # Local ChromaDB persistent store
├── tests/             # Automated pytest suite
├── config.py          # Application configuration
├── seed.py            # Database initialization & seed script
├── run.py             # Flask application entry point
└── requirements.txt   # Python dependencies
```

---

## Setup & Installation

### 1. Prerequisites
- Python 3.10+
- Virtual environment tool (`venv`)

### 2. Installation & Virtual Environment Setup

Navigate to the project root directory:
```bash
cd oes
python -m venv venv
```

Activate the virtual environment:
- **Windows (PowerShell)**: `.\venv\Scripts\Activate.ps1`
- **Linux / macOS**: `source venv/bin/activate`

Install dependencies:
```bash
pip install -r requirements.txt
```

---

## Environment Variables Configuration

You can set environment variables directly in your terminal or configure them in `config.py`.

```bash
# Windows PowerShell
$env:OPEN_ROUTER_API_KEY = "your-openrouter-api-key"
$env:SECRET_KEY = "your-custom-secret-key"

# Linux / macOS
export OPEN_ROUTER_API_KEY="your-openrouter-api-key"
export SECRET_KEY="your-custom-secret-key"
```

---

## Initializing & Seeding Data

Run the seeding script to create database tables, seed mock users (admin and students), create pre-built published exams, and index question embeddings into ChromaDB:

```bash
python seed.py
```

### Pre-configured Login Credentials:
- **Administrator**:
  - Email: `admin@oes.com`
  - Password: `admin123`
- **Student**:
  - Email: `student@oes.com`
  - Password: `student123`

---

## Running the Application Locally

Start the local Flask development server:
```bash
python run.py
```
Open your browser and navigate to:
[http://127.0.0.1:5000](http://127.0.0.1:5000)

---

## Running Automated Tests

Run the pytest test suite to verify route handlers, timer authority, auto-save API, and scoring logic:

```bash
python -m pytest tests/test_oes.py
```
>>>>>>> 2fad3fa (Deploy Online Examination System)
