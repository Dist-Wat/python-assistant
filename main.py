# libraries
import win32com.client
import speech_recognition as sr
import webbrowser
import openai

# initializing
speaker = win32com.client.Dispatch("SAPI.SpVoice")


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-uk")
            print(f"user said: {query}\n")
            return query
        except Exception as e:
            return "some error occurred, try again"


openai.api_key = 'YOUR-api-key'


def chat_with_bot(messages):
    chat_log = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
    )
    return chat_log.choices[0].message

# Upper stuff


def main():
    print("Welcome to the chatbot!")
    speaker.Speak("Welcome to the chatbot!")

    chat_history = []

    # chat-loop
    while True:
        # taking commands
        print("listening...")
        text = take_command()

        # allow opening Websites
        if 'open google' in text.lower():
            speaker.Speak("opening google")
            webbrowser.open("https://www.google.com")

        if 'open youtube' in text.lower():
            speaker.speak("opening youtube")
            webbrowser.open("https://youtube.com")

        if 'open wikipedia' in text.lower():
            speaker.Speak("opening wikipedia")
            webbrowser.open("https://www.wikipedia.org/")

        # handle control
        if text == 'stop':
            break
        if text == 'continue':
            continue

        chat_history.append({
                'role': 'system',
                'content': text
            })

        # Generate a response
        bot_response = chat_with_bot(chat_history)

        print("Chatbot:", bot_response['content'])
        speaker.Speak(bot_response['content'])

        chat_history.append({
                'role': 'system',
                'content': bot_response['content']
            })

        print("Goodbye!")


if __name__ == "__main__":
    main()
