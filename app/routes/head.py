from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.models.schema import db, User, Exam, Attempt, IntegrityEvent
from app.routes.auth import head_required

head_bp = Blueprint('head', __name__, url_prefix='/head')


@head_bp.route('/dashboard')
@login_required
@head_required
def dashboard():
    total_users = User.query.count()
    total_students = User.query.filter_by(role='student').count()
    total_admins = User.query.filter_by(role='admin').count()
    total_heads = User.query.filter_by(role='head').count()

    total_exams = Exam.query.count()
    published_exams = Exam.query.filter_by(is_published=True).count()

    attempts = Attempt.query.all()
    total_attempts = len(attempts)
    graded_attempts = [a for a in attempts if a.status in ['submitted', 'graded']]
    passed_attempts = [a for a in graded_attempts if a.passed]
    pass_rate = (len(passed_attempts) / len(graded_attempts) * 100.0) if graded_attempts else 0.0

    total_integrity_flags = IntegrityEvent.query.count()
    recent_attempts = Attempt.query.order_by(Attempt.start_time.desc()).limit(6).all()
    recent_users = User.query.order_by(User.created_at.desc()).limit(6).all()

    return render_template(
        'head/dashboard.html',
        total_users=total_users,
        total_students=total_students,
        total_admins=total_admins,
        total_heads=total_heads,
        total_exams=total_exams,
        published_exams=published_exams,
        total_attempts=total_attempts,
        pass_rate=pass_rate,
        total_integrity_flags=total_integrity_flags,
        recent_attempts=recent_attempts,
        recent_users=recent_users
    )


@head_bp.route('/users')
@login_required
@head_required
def users():
    role_filter = request.args.get('role', 'all').strip().lower()
    search_q = request.args.get('q', '').strip()

    query = User.query
    if role_filter in ['student', 'admin', 'head']:
        query = query.filter_by(role=role_filter)

    if search_q:
        search_pattern = f"%{search_q}%"
        query = query.filter((User.name.ilike(search_pattern)) | (User.email.ilike(search_pattern)))

    user_list = query.order_by(User.id.asc()).all()

    return render_template(
        'head/users.html',
        users=user_list,
        role_filter=role_filter,
        search_q=search_q,
        total_count=len(user_list)
    )


@head_bp.route('/users/<int:user_id>/delete', methods=['POST'])
@login_required
@head_required
def delete_user(user_id):
    if user_id == current_user.id:
        flash("Safety Protection: You cannot delete your own active Head account.", "danger")
        return redirect(url_for('head.users'))

    user = User.query.get_or_404(user_id)
    name = user.name
    email = user.email
    
    try:
        db.session.delete(user)
        db.session.commit()
        flash(f"Successfully deleted user account: {name} ({email})", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Failed to delete user account: {str(e)}", "danger")

    return redirect(url_for('head.users'))


@head_bp.route('/users/<int:user_id>/role', methods=['POST'])
@login_required
@head_required
def change_user_role(user_id):
    if user_id == current_user.id:
        flash("Safety Protection: You cannot alter your own role.", "warning")
        return redirect(url_for('head.users'))

    new_role = request.form.get('role', '').strip().lower()
    if new_role not in ['student', 'admin', 'head']:
        flash("Invalid role specified.", "danger")
        return redirect(url_for('head.users'))

    user = User.query.get_or_404(user_id)
    old_role = user.role
    user.role = new_role
    db.session.commit()
    flash(f"Updated role for {user.name} from '{old_role}' to '{new_role}'.", "success")
    return redirect(url_for('head.users'))


@head_bp.route('/audit')
@login_required
@head_required
def audit_log():
    events = IntegrityEvent.query.order_by(IntegrityEvent.timestamp.desc()).limit(100).all()
    return render_template('head/audit.html', events=events)
