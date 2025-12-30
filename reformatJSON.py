"""
Parse Crewhu notification emails and extract survey data.

Reads raw email JSON exported from Outlook and extracts structured
survey information (ticket number, rating, feedback) into a clean format.
"""

import json
import re
from pathlib import Path

# ---------------------------------------------------------
# Configuration
# ---------------------------------------------------------
DRY_RUN = False  # Set to True to preview output without writing file

INPUT_FILE = Path("crewhu_notifications_NEW.json")
OUTPUT_FILE = Path("crewhu_surveys_clean.json")


# ---------------------------------------------------------
# Regex Patterns
# ---------------------------------------------------------
# Pattern for "Review" style emails
# "Name from Company gave a Rating rating to Employee for Categories on ticket# ID (Desc)."
REGEX_REVIEW = re.compile(
    r"(?P<customer>.*?) from (?P<company>.*?) gave a (?P<rating>.*?) rating to (?P<employee>.*?) for (?P<categories>.*?) on ticket# (?P<ticket_id>\d+)\s*\((?P<ticket_desc>.*?)\)",
    re.IGNORECASE
)

# Pattern for "Woohoo" style emails
# "Name from Company gave a Rating Rating for Categories on ticket# ID (Desc) to your colleague Employee."
REGEX_WOOHOO = re.compile(
    r"(?P<customer>.*?) from (?P<company>.*?) gave a (?P<rating>.*?) Rating for (?P<categories>.*?) on ticket# (?P<ticket_id>\d+)\s*\((?P<ticket_desc>.*?)\) to your colleague (?P<employee>.*?)",
    re.IGNORECASE
)

# Pattern for Feedback
REGEX_FEEDBACK = re.compile(
    r"Customer feedback:\s*(?:\"(?P<quote>.*?)\"|(?P<none>No feedback provided))",
    re.IGNORECASE | re.DOTALL
)


# ---------------------------------------------------------
# Parsing Logic
# ---------------------------------------------------------
def extract_survey_from_email(email):
    """Extract survey data from a single email. Returns dict or None."""
    subject = email.get('Subject', '')
    full_body = email.get('FullBody', '')

    # Skip if it's not a rating email
    if 'rating' not in subject.lower():
        return None

    match_data = None

    # Split by lines to find the specific sentence and avoid "EXTERNAL EMAIL" headers
    cleaned_lines = [line.strip() for line in full_body.splitlines() if line.strip()]

    for line in cleaned_lines:
        # Try matching the "Review" format line
        m1 = REGEX_REVIEW.search(line)
        if m1:
            match_data = m1.groupdict()
            break

        # Try matching the "Woohoo" format line
        m2 = REGEX_WOOHOO.search(line)
        if m2:
            match_data = m2.groupdict()
            break

    if not match_data:
        return None

    # Extract Ticket ID
    ticket_id_str = match_data['ticket_id']
    ticket_id_int = int(ticket_id_str)

    # Cleanup Extracted Data
    customer = match_data['customer'].strip()
    company = match_data['company'].strip()
    rating = match_data['rating'].strip()
    employee = match_data['employee'].strip().rstrip('.')
    categories = match_data['categories'].strip()
    ticket_desc = match_data['ticket_desc'].strip()

    # Extract Feedback from the full body block
    feedback_text = "No feedback provided."
    fb_match = REGEX_FEEDBACK.search(full_body)
    if fb_match:
        if fb_match.group('quote'):
            feedback_text = fb_match.group('quote').strip()
        elif fb_match.group('none'):
            feedback_text = "No feedback provided."

    # Construct the summary sentence
    summary_sentence = (
        f"{customer} from {company} just gave a {rating} rating to {employee} "
        f"for {categories} on ticket# {ticket_id_str} ({ticket_desc})."
    )

    return {
        "ticket_number": ticket_id_int,
        "summary": summary_sentence,
        "customer_feedback": feedback_text
    }


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------
def main():
    # Check input file exists
    if not INPUT_FILE.exists():
        print(f"ERROR: Could not find {INPUT_FILE}")
        return

    # Load emails
    try:
        with open(INPUT_FILE, 'r', encoding='utf-8-sig') as f:
            emails = json.load(f)
    except json.JSONDecodeError:
        print(f"ERROR: Invalid JSON format in {INPUT_FILE}")
        return

    print(f"Scanning {len(emails)} emails...")

    if DRY_RUN:
        print("!!! DRY RUN MODE - Output file will not be written !!!\n")

    # Process emails and deduplicate by ticket ID
    surveys = {}
    skipped = 0

    for email in emails:
        result = extract_survey_from_email(email)
        if result:
            surveys[result['ticket_number']] = result
        else:
            skipped += 1

    # Sort by ticket number
    sorted_data = sorted(surveys.values(), key=lambda x: x['ticket_number'])

    # Preview or save
    if DRY_RUN:
        print("Preview of extracted data:")
        print("-" * 40)
        for survey in sorted_data[:5]:  # Show first 5
            print(f"Ticket #{survey['ticket_number']}")
            print(f"  {survey['summary'][:60]}...")
            print(f"  Feedback: {survey['customer_feedback'][:40]}...")
            print()
        if len(sorted_data) > 5:
            print(f"... and {len(sorted_data) - 5} more surveys")
    else:
        try:
            with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
                json.dump(sorted_data, f, indent=4)
            print(f"Output saved to: {OUTPUT_FILE}")
        except IOError as e:
            print(f"ERROR saving file: {e}")
            return

    # Summary
    print("\n" + "=" * 40)
    print("SUMMARY")
    print("=" * 40)
    print(f"Total emails scanned: {len(emails)}")
    print(f"Surveys extracted:    {len(sorted_data)}")
    print(f"Emails skipped:       {skipped}")

    if DRY_RUN:
        print("\nDRY RUN COMPLETE - Set DRY_RUN = False to write output file.")


if __name__ == "__main__":
    main()
