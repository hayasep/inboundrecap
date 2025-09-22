# Backstock Filler Emailer

Flask app that renders an Excel template in the browser (Luckysheet),
lets you edit it, then **downloads or emails** the filled workbook.

## Local setup
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env  # fill values
python app.py
# open http://localhost:5000
