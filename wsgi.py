import sys
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Add project directory to path
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Change to project directory
os.chdir(current_dir)

# Import app
from app import app as application, db
from seed import seed_database

# Initialize database
try:
    with application.app_context():
        db.create_all()
        seed_database()
except Exception as e:
    print(f"Database init error: {e}")

# For Gunicorn
app = application
