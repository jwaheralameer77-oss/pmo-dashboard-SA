import os
import sys

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from seed import seed_database

# Initialize database
with app.app_context():
    try:
        db.create_all()
        seed_database()
        print("Database initialized successfully")
    except Exception as e:
        print(f"Database initialization warning: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
