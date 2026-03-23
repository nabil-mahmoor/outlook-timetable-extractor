import os
from datetime import datetime
from auth import get_access_token
from pdf_handler import find_timetable_page, save_page_as_image
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
                date_str = str(timestamp).split(" ")[0]
                pdf_path = download_pdf(attachment, date_str)

                page_index = find_timetable_page(pdf_path)
                if not page_index:
                    print("Could not find your timetable page. Check the file for spelling errors")
                else:
                    save_page_as_image(pdf_path, page_index, date_str)
                    
                    # Clean up the temporary PDF
                    os.remove(pdf_path)
                    print("Temporary PDF removed.")

    except Exception as e:
        print(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
