import sys
import os

# Ensure fresh imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Clear any cached modules
modules_to_clear = [k for k in sys.modules.keys() if k.startswith('app') or k.startswith('models')]
for m in modules_to_clear:
    del sys.modules[m]

# Now import fresh
import app

# Run
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.app.run(debug=False, host='0.0.0.0', port=port, use_reloader=False)
