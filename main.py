from auth import get_access_token
from mail import get_latest_timetable_email

def main():
    try:
        token = get_access_token()
        print("Access token aquired successfully")
        
        email = get_latest_timetable_email(token)
        if email:
            print(f"Found email: '{email.get('subject')}' received on {email.get('receivedDateTime')}")
        else:
            print("No timetable email found.")

    except Exception as e:
        print(f"Something went wrong: {e}")


if __name__ == "__main__":
    main()
