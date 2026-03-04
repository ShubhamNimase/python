import time
import pandas as pd
import win32com.client
import re

def insert_html_body(original_html, insert_text):
    """Inserts new HTML into the existing HTML body, ideally after the <body> tag."""
    body_match = re.search(r'(<body[^>]*>)', original_html, re.IGNORECASE)
    if body_match:
        # Insert right after the <body> tag
        idx = body_match.end()
        return original_html[:idx] + insert_text + original_html[idx:]
    else:
        # If no body tag found, just prepend
        return insert_text + original_html

def send_followups(excel_filepath, new_subject="Follow-up"):
    # Load data from Excel
    print(f"Loading data from {excel_filepath}...")
    try:
        df = pd.read_excel(excel_filepath)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Initialize Outlook application
    print("Connecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # 5 refers to the Sent Items folder
        sent_folder = namespace.GetDefaultFolder(5)
        sent_items = sent_folder.Items

        # Sort items by SentOn ascending to find the earliest/first email sent
        sent_items.Sort("[SentOn]", False)
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return

    # Iterate over the rows in the Excel file
    total_emails = len(df)
    for index, row in df.iterrows():
        # Ensure we read the correct columns 'Name' and 'Email ID'
        full_name = str(row.get('Name', '')).strip()
        email_id = str(row.get('Email ID', '')).strip()

        if not full_name or not email_id or full_name.lower() == 'nan' or email_id.lower() == 'nan':
            print(f"Skipping row {index + 1}: Missing 'Name' or 'Email ID'.")
            continue

        # Extract first name
        first_name = full_name.split()[0]

        print(f"\nProcessing follow-up for {full_name} ({email_id})...")

        found_email = None

        try:
            # Use Restrict to efficiently find emails sent to this email address
            # The filter format requires single quotes around the search string
            # Also, we check if the recipient address contains the email_id
            filter_str = f"@SQL=\"urn:schemas:httpmail:displayto\" LIKE '%{email_id}%' OR \"urn:schemas:httpmail:to\" LIKE '%{email_id}%'"
            filtered_items = sent_items.Restrict(filter_str)

            # Sort the restricted items to ensure we get the first one sent
            filtered_items.Sort("[SentOn]", False)

            # Additional manual check because the restriction might match loosely
            for item in filtered_items:
                if getattr(item, 'Class', 0) == 43: # MailItem
                    try:
                        for rec in item.Recipients:
                            rec_address = getattr(rec, 'Address', '').lower()
                            rec_name = getattr(rec, 'Name', '').lower()
                            if email_id.lower() in rec_address or email_id.lower() in rec_name:
                                found_email = item
                                break
                    except Exception:
                        pass

                if found_email:
                    break
        except Exception as e:
            print(f"Error searching for email: {e}")

        if found_email:
            try:
                # Reply All to the found email
                reply = found_email.ReplyAll()

                # Replace the old subject with the new subject
                reply.Subject = new_subject

                # Follow-up message using the first name
                followup_text = (
                    f"<p>Hi {first_name},</p>"
                    f"<p>I am following up on the email below. "
                    f"Please let me know if you have any questions or updates.</p>"
                    f"<p>Best regards</p>"
                )

                # Correctly insert the follow-up text into the existing HTML body
                reply.HTMLBody = insert_html_body(reply.HTMLBody, followup_text)

                # Send the email
                reply.Send()
                print(f"Successfully sent follow-up to {first_name} ({email_id}).")

                # Add a time gap of 30 seconds between 2 consecutive emails
                # We only wait if this is not the very last email to process
                if index < total_emails - 1:
                    print("Waiting 30 seconds before sending the next email...")
                    time.sleep(30)

            except Exception as e:
                print(f"Error sending email to {email_id}: {e}")
        else:
            print(f"Could not find a sent email for {email_id} in the Sent Items.")

if __name__ == "__main__":
    # Ensure you have 'contacts.xlsx' in the same directory, or provide the full path
    excel_file = "contacts.xlsx"
    # Set the desired new subject line here
    new_subject_line = "Follow-up"
    send_followups(excel_file, new_subject_line)
