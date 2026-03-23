from datetime import datetime
from auth import get_access_token
from mail import get_latest_timetable_email, get_pdf_attachment, download_pdf

def main():
    try:
        token = get_access_token()
        print("Access token aquired successfully")
        
        email = get_latest_timetable_email(token)

        if not email:
            print("No timetable email found.")
        else:
            timestamp = datetime.strptime(email.get('receivedDateTime'), "%Y-%m-%dT%H:%M:%SZ")
            print(f"Found email: '{email.get('subject')}' received on {timestamp}")

            attachment = get_pdf_attachment(token, email.get("id"))

            if not attachment:
                print("No PDF attachment found in email.")
            else:
                date = str(timestamp).split(" ")[0]
                download_pdf(attachment, date)

    except Exception as e:
        print(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
