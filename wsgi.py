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
from app import app as application, init_db

# Initialize database on first load
try:
    init_db()
except:
    pass

# For Gunicorn
app = application
