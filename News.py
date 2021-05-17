import requests
from win32com.client import Dispatch
import json
import time

input_from_user = "yes"
try:
    while input_from_user == "yes":
        def news_paper(str1):
            speak = Dispatch("SAPI.Spvoice")
            speak.Speak(str1)


        categories = int(input("1-business\n2-entertainment\n3-general\n4-science\n5-Top "
                               "Headlines\n6-technology\n7-health\n"))
        country = input("Enter your country,Example {in-India}{at-australia}{us-United states}\n ")
        api_key = "62b0f85f88014ebcbd167894b91707f3"
        if __name__ == '__main__':
            if categories == 7:
                url = f"https://newsapi.org/v2/top-headlines?country={country}&category=health&apiKey" \
                      f"={api_key}"
            if categories == 6:
                url = f"https://newsapi.org/v2/top-headlines?country={country}&category=technology&apiKey" \
                      f"={api_key}"
            if categories == 5:
                url = f"https://newsapi.org/v2/top-headlines?country={country}&apiKey" \
                      f"={api_key}"
            if categories == 4:
                sources = input()
                url = f"https://newsapi.org/v2/top-headlines?country={country}&category=science&apiKey" \
                      f"={api_key} "
            if categories == 3:
                url = f"https://newsapi.org/v2/top-headlines?country={country}&category=general&apiKey" \
                      f"={api_key} "
            if categories == 2:
                url = f"https://newsapi.org/v2/top-headlines?country={country}&category=entertainment&apiKey" \
                      f"={api_key} "
            if categories == 1:
                url = f"https://newsapi.org/v2/top-headlines?country={country}&category=business&apiKey" \
                      f"={api_key} "
            response = requests.get(url)
            text = response.text
            my_json = json.loads(text)
            news_input = int(input("How many News you want , news can't be more than 10\n"))
            input_from_user = input("what you want to do read or listen the news or both\n")
            for i in range(news_input):
                time.sleep(2)
                if i == 0:
                    news_paper("Our first news is")
                if i == 1:
                    news_paper("Our second news is")
                if i == 2:
                    news_paper("Our third news is")
                if i == 3:
                    news_paper("Our fourth news is")
                if i == 4:
                    news_paper("Our fifth news is")
                if i == 5:
                    news_paper("Our sixth news is")
                if i == 6:
                    news_paper("Our seventh news is")
                if i == 7:
                    news_paper("Our eight news is")
                if i == 8:
                    news_paper("Our ninth news is")
                if i == 9:
                    news_paper("Our Tenth news is")
                if news_input >= 10:
                    print("Not Allowed")
                    break
                if input_from_user == "read":
                    print("for more information about it visit", [my_json['articles'][i]['url']])
                    print(my_json['articles'][i]['title'])
                    print("\n")
                if input_from_user == "listen":
                    print("for more information about it visit", [my_json['articles'][i]['url']])
                    news_paper(my_json['articles'][i]['title'])
                    print("\n")
                if input_from_user == "both":
                    print("for more information about it visit", [my_json['articles'][i]['url']])
                    news_paper(my_json['articles'][i]['title'])
                    print(my_json['articles'][i]['title'])
                    print("\n")
except Exception as error:
    print(error)
