#!/bin/bash
#!/bin/bash
export FLASK_APP=app.py
export FLASK_ENV=production

# Railway automatically injects $PORT as an env variable
# Use Python to read it safely
python -m flask run --host=0.0.0.0 --port=${PORT:-8000}
