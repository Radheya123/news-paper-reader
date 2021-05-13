import requests
from win32com.client import Dispatch
import json
input_from_user = "yes"
try:
    while input_from_user == "yes":
        def news_paper(str1):
            speak = Dispatch("SAPI.Spvoice")
            speak.Speak(str1)


        if __name__ == '__main__':
            url = ('https://newsapi.org/v2/top-headlines?'
                   'sources=bbc-sport&'
                   'apiKey=49e391e7066c4158937096fb5e55fb5d')
            response = requests.get(url)
            text = response.text
            my_json = json.loads(text)
            news_input = int(input("How many News you want to read news cant be more than 10\n"))
            input_from_user = input("what you want to do read or listen the news or nothing\n")
            if news_input > 10:
                print("news not allowed more than 10")
                break
            for i in range(news_input):
                if input_from_user == "read":
                    print(my_json['articles'][i]['title'])
                if input_from_user == "listen":
                    news_paper(my_json['articles'][i]['title'])
                if input_from_user == "both":
                    print(my_json['articles'][i]['title'])
                    news_paper(my_json['articles'][i]['title'])
                # else:
                #     print("not available")

except IndexError as error:
    print(error)
