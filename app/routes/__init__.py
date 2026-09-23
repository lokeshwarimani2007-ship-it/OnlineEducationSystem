from .auth import auth_bp
from .student import student_bp
from .admin import admin_bp
from .api import api_bp

__all__ = ['auth_bp', 'student_bp', 'admin_bp', 'api_bp']
