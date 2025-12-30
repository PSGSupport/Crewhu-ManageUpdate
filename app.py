"""
Flask web application for Crewhu-ConnectWise integration.

Provides a web interface for running the automation scripts with
real-time status updates and file upload capabilities.
"""

import os
import json
import queue
import threading
from pathlib import Path
from datetime import datetime, timezone
from flask import Flask, render_template, request, jsonify, Response
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = Path('uploads')

# Ensure upload folder exists
app.config['UPLOAD_FOLDER'].mkdir(exist_ok=True)

# Global status queue for SSE
status_queues = {}


# ==========================================
# STATUS STREAMING (SSE)
# ==========================================
def get_status_queue(session_id):
    """Get or create a status queue for a session."""
    if session_id not in status_queues:
        status_queues[session_id] = queue.Queue()
    return status_queues[session_id]


def send_status(session_id, message, progress=None, status_type="info"):
    """Send a status update to the client."""
    if session_id in status_queues:
        status_queues[session_id].put({
            "message": message,
            "progress": progress,
            "type": status_type,
            "timestamp": datetime.now().isoformat()
        })


def cleanup_queue(session_id):
    """Clean up a session's queue."""
    if session_id in status_queues:
        del status_queues[session_id]


# ==========================================
# IMPORT SCRIPTS (with status callbacks)
# ==========================================
def run_reformat_json(input_file, output_file, session_id, dry_run=False):
    """Run the reformatJSON logic with status updates."""
    import re

    send_status(session_id, "Starting JSON parsing...", 0)

    # Regex patterns from reformatJSON.py
    REGEX_REVIEW = re.compile(
        r"(?P<customer>.*?) from (?P<company>.*?) gave a (?P<rating>.*?) rating to (?P<employee>.*?) for (?P<categories>.*?) on ticket# (?P<ticket_id>\d+)\s*\((?P<ticket_desc>.*?)\)",
        re.IGNORECASE
    )
    REGEX_WOOHOO = re.compile(
        r"(?P<customer>.*?) from (?P<company>.*?) gave a (?P<rating>.*?) Rating for (?P<categories>.*?) on ticket# (?P<ticket_id>\d+)\s*\((?P<ticket_desc>.*?)\) to your colleague (?P<employee>.+?)\.?$",
        re.IGNORECASE
    )
    REGEX_FEEDBACK = re.compile(
        r"Customer feedback:\s*(?:\"(?P<quote>.*?)\"|(?P<none>No feedback provided))",
        re.IGNORECASE | re.DOTALL
    )

    try:
        with open(input_file, 'r', encoding='utf-8-sig') as f:
            emails = json.load(f)
    except Exception as e:
        send_status(session_id, f"Error reading file: {e}", 100, "error")
        return {"success": False, "error": str(e)}

    send_status(session_id, f"Loaded {len(emails)} emails", 10)

    surveys = {}
    skipped = 0

    for i, email in enumerate(emails):
        progress = 10 + int((i / len(emails)) * 80)

        subject = email.get('Subject', '') or email.get('summary', '')
        full_body = email.get('FullBody', '') or email.get('full_clean_body', '')

        if 'rating' not in subject.lower() and 'gave a' not in full_body.lower():
            skipped += 1
            continue

        match_data = None
        cleaned_lines = [line.strip() for line in full_body.splitlines() if line.strip()]

        for line in cleaned_lines:
            m1 = REGEX_REVIEW.search(line)
            if m1:
                match_data = m1.groupdict()
                break
            m2 = REGEX_WOOHOO.search(line)
            if m2:
                match_data = m2.groupdict()
                break

        if not match_data:
            skipped += 1
            continue

        ticket_id_int = int(match_data['ticket_id'])

        feedback_text = "No feedback provided."
        fb_match = REGEX_FEEDBACK.search(full_body)
        if fb_match:
            if fb_match.group('quote'):
                feedback_text = fb_match.group('quote').strip()

        summary_sentence = (
            f"{match_data['customer'].strip()} from {match_data['company'].strip()} "
            f"just gave a {match_data['rating'].strip()} rating to {match_data['employee'].strip().rstrip('.')} "
            f"for {match_data['categories'].strip()} on ticket# {match_data['ticket_id']} "
            f"({match_data['ticket_desc'].strip()})."
        )

        surveys[ticket_id_int] = {
            "ticket_number": ticket_id_int,
            "summary": summary_sentence,
            "customer_feedback": feedback_text
        }

        if i % 10 == 0:
            send_status(session_id, f"Processing email {i+1}/{len(emails)}...", progress)

    sorted_data = sorted(surveys.values(), key=lambda x: x['ticket_number'])

    if not dry_run:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(sorted_data, f, indent=4)

    send_status(session_id, f"Parsing complete! Extracted {len(sorted_data)} surveys, skipped {skipped}", 100, "success")

    return {
        "success": True,
        "total_emails": len(emails),
        "surveys_extracted": len(sorted_data),
        "skipped": skipped,
        "dry_run": dry_run
    }


def run_post_ratings(csv_file, session_id, dry_run=True):
    """Run the POST_Ratings logic with status updates."""
    import csv
    import requests
    from connectwise_api import validate_credentials, get_headers, tickets_url

    send_status(session_id, "Starting ratings update...", 0)

    is_valid, missing = validate_credentials()
    if not is_valid:
        send_status(session_id, f"Missing credentials: {', '.join(missing)}", 100, "error")
        return {"success": False, "error": f"Missing credentials: {', '.join(missing)}"}

    RATING_MAP = {
        "AWESOME": "Positive",
        "POSITIVE": "Positive",
        "NEGATIVE": "Negative",
    }

    try:
        with open(csv_file, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
    except Exception as e:
        send_status(session_id, f"Error reading CSV: {e}", 100, "error")
        return {"success": False, "error": str(e)}

    send_status(session_id, f"Loaded {len(rows)} rows from CSV", 10)

    headers = get_headers()
    updated = 0
    skipped = 0
    errors = 0

    for i, row in enumerate(rows):
        progress = 10 + int((i / len(rows)) * 85)

        ticket_id = str(row.get('Ticket#', '')).strip()
        if '.' in ticket_id:
            ticket_id = ticket_id.split('.')[0]

        if not ticket_id or not ticket_id.isdigit():
            skipped += 1
            continue

        rating = str(row.get('Rating', '')).strip().upper()
        if rating not in RATING_MAP:
            skipped += 1
            continue

        rating_value = RATING_MAP[rating]

        if dry_run:
            send_status(session_id, f"[DRY RUN] Ticket {ticket_id}: Would set rating to {rating_value}", progress)
            updated += 1
        else:
            url = tickets_url(ticket_id)
            payload = [{
                "op": "replace",
                "path": "customFields",
                "value": [{
                    "id": 18,
                    "caption": "Latest Crewhu Rating",
                    "type": "Text",
                    "entryMethod": "EntryField",
                    "numberOfDecimals": 0,
                    "value": rating_value
                }]
            }]

            try:
                response = requests.patch(url, headers=headers, json=payload)
                if response.status_code == 200:
                    send_status(session_id, f"Ticket {ticket_id}: Updated rating to {rating_value}", progress)
                    updated += 1
                else:
                    send_status(session_id, f"Ticket {ticket_id}: Error {response.status_code}", progress, "warning")
                    errors += 1
            except Exception as e:
                send_status(session_id, f"Ticket {ticket_id}: Network error", progress, "error")
                errors += 1

    send_status(session_id, f"Ratings update complete! Updated: {updated}, Skipped: {skipped}, Errors: {errors}", 100, "success")

    return {
        "success": True,
        "total_rows": len(rows),
        "updated": updated,
        "skipped": skipped,
        "errors": errors,
        "dry_run": dry_run
    }


def run_post_notes(json_file, session_id, dry_run=True):
    """Run the POST_Notes_Internal logic with status updates."""
    import requests
    from connectwise_api import validate_credentials, get_headers, ticket_notes_url

    send_status(session_id, "Starting notes posting...", 0)

    is_valid, missing = validate_credentials()
    if not is_valid:
        send_status(session_id, f"Missing credentials: {', '.join(missing)}", 100, "error")
        return {"success": False, "error": f"Missing credentials: {', '.join(missing)}"}

    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            parsed_data = json.load(f)
    except Exception as e:
        send_status(session_id, f"Error reading JSON: {e}", 100, "error")
        return {"success": False, "error": str(e)}

    send_status(session_id, f"Loaded {len(parsed_data)} survey entries", 10)

    headers = get_headers()
    posted = 0
    errors = 0

    for i, entry in enumerate(parsed_data):
        progress = 10 + int((i / len(parsed_data)) * 85)

        ticket_id = entry.get("ticket_number")
        summary = entry.get("summary", "").strip()
        feedback = entry.get("customer_feedback", "No feedback provided.").strip()

        if not ticket_id:
            continue

        final_note_text = f"{summary}\n\nCustomer feedback:\n{feedback}"

        if dry_run:
            send_status(session_id, f"[DRY RUN] Ticket {ticket_id}: Would post internal note", progress)
            posted += 1
        else:
            # Delete old automated notes first
            try:
                notes_url = ticket_notes_url(ticket_id)
                resp = requests.get(notes_url, headers=headers)
                if resp.status_code == 200:
                    for note in resp.json():
                        text = (note.get("text") or "")
                        if "just gave a" in text and "Customer feedback:" in text:
                            del_url = ticket_notes_url(ticket_id, note.get("id"))
                            requests.delete(del_url, headers=headers)
            except:
                pass

            # Post new note
            payload = {
                "text": final_note_text,
                "detailDescriptionFlag": False,
                "internalAnalysisFlag": True,
                "resolutionFlag": False,
                "createdBy": "Crewhu API",
                "dateCreated": datetime.now(timezone.utc).isoformat()
            }

            try:
                url = ticket_notes_url(ticket_id)
                response = requests.post(url, headers=headers, json=payload)
                if response.status_code == 201:
                    send_status(session_id, f"Ticket {ticket_id}: Posted internal note", progress)
                    posted += 1
                else:
                    send_status(session_id, f"Ticket {ticket_id}: Error {response.status_code}", progress, "warning")
                    errors += 1
            except Exception as e:
                send_status(session_id, f"Ticket {ticket_id}: Network error", progress, "error")
                errors += 1

    send_status(session_id, f"Notes posting complete! Posted: {posted}, Errors: {errors}", 100, "success")

    return {
        "success": True,
        "total_entries": len(parsed_data),
        "posted": posted,
        "errors": errors,
        "dry_run": dry_run
    }


# ==========================================
# ROUTES
# ==========================================
@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')


@app.route('/api/status')
def check_status():
    """Check API credentials status."""
    from connectwise_api import validate_credentials
    is_valid, missing = validate_credentials()
    return jsonify({
        "credentials_valid": is_valid,
        "missing": missing
    })


@app.route('/api/stream/<session_id>')
def stream(session_id):
    """SSE endpoint for status updates."""
    def generate():
        q = get_status_queue(session_id)
        while True:
            try:
                status = q.get(timeout=30)
                if status.get("type") == "done":
                    yield f"data: {json.dumps(status)}\n\n"
                    break
                yield f"data: {json.dumps(status)}\n\n"
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'keepalive'})}\n\n"

    return Response(generate(), mimetype='text/event-stream')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload."""
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    filename = secure_filename(file.filename)
    filepath = app.config['UPLOAD_FOLDER'] / filename
    file.save(filepath)

    return jsonify({
        "success": True,
        "filename": filename,
        "path": str(filepath)
    })


@app.route('/api/parse-emails', methods=['POST'])
def parse_emails():
    """Parse Crewhu notification emails."""
    data = request.json
    session_id = data.get('session_id', 'default')
    input_file = data.get('input_file')
    dry_run = data.get('dry_run', True)

    if not input_file:
        return jsonify({"error": "No input file specified"}), 400

    input_path = app.config['UPLOAD_FOLDER'] / input_file
    output_path = app.config['UPLOAD_FOLDER'] / 'crewhu_surveys_clean.json'

    def run_task():
        result = run_reformat_json(input_path, output_path, session_id, dry_run)
        send_status(session_id, json.dumps(result), 100, "done")

    thread = threading.Thread(target=run_task)
    thread.start()

    return jsonify({"status": "started", "session_id": session_id})


@app.route('/api/post-ratings', methods=['POST'])
def post_ratings():
    """Update ticket ratings from CSV."""
    data = request.json
    session_id = data.get('session_id', 'default')
    csv_file = data.get('csv_file')
    dry_run = data.get('dry_run', True)

    if not csv_file:
        return jsonify({"error": "No CSV file specified"}), 400

    csv_path = app.config['UPLOAD_FOLDER'] / csv_file

    def run_task():
        result = run_post_ratings(csv_path, session_id, dry_run)
        send_status(session_id, json.dumps(result), 100, "done")

    thread = threading.Thread(target=run_task)
    thread.start()

    return jsonify({"status": "started", "session_id": session_id})


@app.route('/api/post-notes', methods=['POST'])
def post_notes():
    """Post internal notes to tickets."""
    data = request.json
    session_id = data.get('session_id', 'default')
    json_file = data.get('json_file', 'crewhu_surveys_clean.json')
    dry_run = data.get('dry_run', True)

    json_path = app.config['UPLOAD_FOLDER'] / json_file

    def run_task():
        result = run_post_notes(json_path, session_id, dry_run)
        send_status(session_id, json.dumps(result), 100, "done")

    thread = threading.Thread(target=run_task)
    thread.start()

    return jsonify({"status": "started", "session_id": session_id})


@app.route('/api/files')
def list_files():
    """List uploaded files."""
    files = []
    for f in app.config['UPLOAD_FOLDER'].iterdir():
        if f.is_file():
            files.append({
                "name": f.name,
                "size": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })
    return jsonify(files)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
