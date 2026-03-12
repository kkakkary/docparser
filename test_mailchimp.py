"""
Test script: parse .docx attachments from the latest matching Gmail and
subscribe/tag each contact in Mailchimp. No CSV or notification emails sent.

Usage:
    python test_mailchimp.py
    python test_mailchimp.py <gmail_message_id>   # target a specific email
"""

import sys
from main import (
    authenticate,
    get_all_parts,
    client_prefix,
    download_attachment,
    extract_docx_text,
    parse_client_info,
    mailchimp_subscribe,
    mailchimp_add_tag,
    GMAIL_QUERY,
)


def main():
    service = authenticate()
    print("Authenticated.\n")

    # Use a provided message ID or fall back to the latest matching email
    if len(sys.argv) > 1:
        msg_id = sys.argv[1]
        print(f"Using provided message ID: {msg_id}\n")
    else:
        results = service.users().messages().list(
            userId='me', q=GMAIL_QUERY, maxResults=1
        ).execute()
        messages = results.get('messages', [])
        if not messages:
            print("No matching emails found.")
            return
        msg_id = messages[0]['id']
        print(f"Using latest matching email (id: {msg_id})\n")

    # Print subject/from/date
    meta = service.users().messages().get(
        userId='me', id=msg_id, format='metadata',
        metadataHeaders=['Subject', 'Date', 'From']
    ).execute()
    headers = {h['name']: h['value'] for h in meta['payload']['headers']}
    print(f"Subject: {headers.get('Subject', 'No Subject')}")
    print(f"From:    {headers.get('From', 'Unknown')}")
    print(f"Date:    {headers.get('Date', '')}\n")

    full_msg = service.users().messages().get(userId='me', id=msg_id).execute()
    parts = get_all_parts(full_msg.get('payload', {}))

    pdf_clients = {
        client_prefix(p['filename'])
        for p in parts
        if p.get('filename', '').lower().endswith('.pdf')
    }
    if pdf_clients:
        print(f"Paired clients (will still be processed for Mailchimp): {pdf_clients}\n")

    for part in parts:
        filename = part.get('filename')
        if not filename or not filename.lower().endswith('.docx'):
            continue

        is_paired = client_prefix(filename) in pdf_clients
        print(f"Processing: {filename}" + (" (paired)" if is_paired else ""))

        data = download_attachment(service, msg_id, part)
        if data is None:
            print(f"  Could not download: {filename}")
            continue

        print(f"  Downloaded ({len(data):,} bytes), parsing with Gemini...")
        text = extract_docx_text(data)
        info = parse_client_info(text, filename)
        if info:
            print(f"  Parsed: {info}")
            if mailchimp_subscribe(info):
                mailchimp_add_tag(info)
        print()

    print("--- Test complete. No CSV or notification emails were sent. ---")


if __name__ == '__main__':
    main()
