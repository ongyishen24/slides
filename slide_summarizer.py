import os

import openai
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/presentations.readonly"]

# The ID of a sample presentation.
PRESENTATION_ID = "1hqroUnCm9iuTxPIeFTVxMCXPAjHZ9C6FXlIf4oMvww0"


def main():
    """Shows basic usage of the Slides API.
    Prints the number of slides and elements in a sample presentation.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("slides", "v1", credentials=creds)

        # Call the Slides API
        presentation = (
            service.presentations().get(presentationId=PRESENTATION_ID).execute()
        )

        slides = presentation.get("slides")
        slide_text = parse_presentation(slides[4:9])

        openai.api_key = os.getenv("OPENAI_API_KEY")

        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Summarize this text for me"},
                {"role": "user", "content": slide_text},
            ],
        )

        print("Slide text: " + slide_text)
        print("OPENAI summary: " + response.choices[0]["message"]["content"])

    except HttpError as err:
        print(err)


def parse_presentation(slides: list) -> str:
    """Returns the text from a list of slides"""
    text_combined = ""

    for slide in slides:
        pageElements = slide.get("pageElements")
        for element in pageElements:
            if element.get("shape") is not None:
                text = element.get("shape").get("text")
                if text is not None:
                    for textElement in text.get("textElements"):
                        if textElement is not None:
                            textRun = textElement.get("textRun")
                            if textRun is not None:
                                text_combined += textRun.get("content")
    return text_combined


if __name__ == "__main__":
    main()
