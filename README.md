# Online Exam Portal

A lightweight Flask + Bootstrap + JavaScript exam portal backed by Excel workbooks.

## Features
- Role-based login and registration
- Admin question bank management
- Exam creation with question selection and publishing
- Timed student exam attempts with auto-submit
- Instant scoring and result storage in Excel

## Run Locally
1. Create a virtual environment.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Start the app:
   ```bash
   python app.py
   ```
4. Open http://127.0.0.1:10000

## Render Deployment
This project is configured to run on Render with `gunicorn`.

Start command:
```bash
gunicorn app:app --bind 0.0.0.0:$PORT
```

The Flask entrypoint also defaults to port `10000` when `PORT` is not provided.

## Demo Accounts
- Admin: admin@example.com / admin123
- Student: student@example.com / student123
