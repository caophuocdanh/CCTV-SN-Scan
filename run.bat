@echo off
pip install -r requirements.txt
ECHO Activating virtual environment...
CALL .\venv\Scripts\activate

ECHO Starting Flask server...
ECHO You can access it from other devices on the same network.

flask run --host=0.0.0.0 --port=5005

PAUSE
