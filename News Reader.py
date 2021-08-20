def speak(str):

    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(str)

print("Catagories:-\n 1. Business\n 2. Entertainment\n 3. Health\n 4. Science\n 5. Sports\n 6. Technology")

if __name__ == '__main__':
    speak("Welcome!,Mr. Lokesh, Which type of News You want")
    speak("Enter the type of news you want from these catagories")

x= input ("Enter your choice:- ")

if x == "Business" or x == "business":

    if __name__ == '__main__':
        speak("Top 5 Business News of today are")
        import requests
        import json

        url = ('https://newsapi.org/v2/top-headlines?'
               'country=in&category=business&'
               'apiKey=e94679063d7046f3ba41579c3fbbd6ca')

        respose = requests.get(url)
        text = respose.text
        my_json = json.loads(text)
        for i in range(0, 6):
            speak(my_json['articles'][i]['title'])


elif x == "Entertainment" or x == "entertainment":
    if __name__ == '__main__':
        speak("Top 5 Entertainment News are")
        import requests
        import json

        url = ('https://newsapi.org/v2/top-headlines?'
               'country=in&category=entertainment&'
               'apiKey=e94679063d7046f3ba41579c3fbbd6ca')

        respose = requests.get(url)
        text = respose.text
        my_json = json.loads(text)
        for i in range(0, 6):
            speak(my_json['articles'][i]['title'])

elif x == "Health" or x == "health":

    if __name__ == '__main__':
        speak("Top 5 Health News are")
        import requests
        import json

        url = ('https://newsapi.org/v2/top-headlines?'
               'country=in&category=health&'
               'apiKey=e94679063d7046f3ba41579c3fbbd6ca')

        respose = requests.get(url)
        text = respose.text
        my_json = json.loads(text)
        for i in range(0, 3):
            speak(my_json['articles'][i]['title'])

elif x == "Science" or x == "science":
    if __name__ == '__main__':
        speak("Top 5 Science News are")
        import requests
        import json

        url = ('https://newsapi.org/v2/top-headlines?'
               'country=in&category=science&'
               'apiKey=e94679063d7046f3ba41579c3fbbd6ca')

        respose = requests.get(url)
        text = respose.text
        my_json = json.loads(text)
        for i in range(0, 5):
            speak(my_json['articles'][i]['title'])

elif x == "Sports" or x == "sports":
    if __name__ == '__main__':
        speak("Top 5 Sports News are")
        import requests
        import json

        url = ('https://newsapi.org/v2/top-headlines?'
               'country=in&category=sports&'
               'apiKey=e94679063d7046f3ba41579c3fbbd6ca')

        respose = requests.get(url)
        text = respose.text
        my_json = json.loads(text)
        for i in range(0, 6):
            speak(my_json['articles'][i]['title'])


elif x == "Technology" or x == "technology":
    if __name__ == '__main__':
        speak("Top 5 Technology News are")
        import requests
        import json

        url = ('https://newsapi.org/v2/top-headlines?'
               'country=in&category=technology&'
               'apiKey=e94679063d7046f3ba41579c3fbbd6ca')

        respose = requests.get(url)
        text = respose.text
        my_json = json.loads(text)
        for i in range(0, 6):
            speak(my_json['articles'][i]['title'])

else:
    if __name__ == '__main__':
        speak("You have entered wrong input, try again")