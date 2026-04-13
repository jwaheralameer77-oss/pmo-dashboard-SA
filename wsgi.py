import sys
import os

# Add project directory to path
path = '/home/yourusername/pmo_app'
if path not in sys.path:
    sys.path.insert(0, path)

# Change to project directory
os.chdir(path)

# Import app
from app import app as application, init_db

# Initialize database on first load
init_db()
