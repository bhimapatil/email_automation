import imaplib
import email
import os
from datetime import datetime, timedelta
from tqdm import tqdm
from dotenv import load_dotenv


load_dotenv()

def download_attachments(msg, folder_path):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = "Mail_Attachment_" + part.get_filename()
        if filename:
            filepath = os.path.join(folder_path, filename)
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))

def main():
    try:
        user = os.getenv("EMAIL")
        password = os.getenv("PASSWORD")
        imap_url = 'imap.gmail.com'

        # Connect to GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using your credentials
        my_mail.login(user, password)

        # Select the Inbox to fetch messages
        my_mail.select('Inbox')

        # Calculate the date range (from yesterday to today)
        today = datetime.now()
        yesterday = today - timedelta(days=2)
        yesterday_date_string = yesterday.strftime("%Y-%m-%d")

        search_query = f'(X-GM-RAW "has:attachment after:{yesterday_date_string}")'
        result, data = my_mail.search(None, search_query)
        email_ids = data[0].split()

        for email_id in tqdm(email_ids, desc="Downloading Attachments"):
            result, email_data = my_mail.fetch(email_id, "(RFC822)")
            raw_email = email_data[0][1]

            msg = email.message_from_bytes(raw_email)

            # Extract 'To', 'From', and 'CC' values
            to_address = msg.get("To", " ")
            from_address = msg.get("From", " ")
            # cc_address = msg.get("CC", "")

            # Define the desired 'To', 'From', and 'CC' addresses
            desired_to_address = os.getenv("TO_ADDRESS")
            desired_from_address = os.getenv("FROM_ADDRESS")
            # desired_cc_address = "desired_cc@example.com"

            # Check if the email matches the desired addresses
            if desired_to_address in to_address and desired_from_address in from_address: #and desired_cc_address in cc_address:
                folder_name = today.strftime(f"Mail_Attachments_%Y-%m-%d")
                folder_path = os.path.abspath(folder_name)

                # Create a folder for each email
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path, mode=0o777)

                # Download attachments for the matching email
                download_attachments(msg, folder_path)

        my_mail.logout()
    except Exception as e:
        print("We are facing temporary server port access denial; please 'Re-run' the program.")



if __name__ == "__main__":
    main()
