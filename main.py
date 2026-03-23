from auth import get_access_token

def main():
    try:
        token = get_access_token()
        print("Access token aquired successfully")
        print(token)
    except Exception as e:
        print(f"Authentication failed: {e}")


if __name__ == "__main__":
    main()
