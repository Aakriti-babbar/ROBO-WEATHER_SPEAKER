import requests
import json
import win32com.client as wincom
speak= wincom.Dispatch("SAPI.SpVoice")
while(True):
    city=input("enter the city: ")
    query=input("What do you want to know about the entered city (temperature/wind condition/pressure/humidity/cloud/air quality/local time):")
    url=f"http://api.weatherapi.com/v1/current.json?key=65ea268716eb4163be6171538242401&q={city}&aqi=yes"
    r=requests.get(url)
    wdic = json.loads(r.text)
    match query :
        case "temperature":
            text=(f"The current temperature in {city} is", wdic["current"]["temp_c"], "degrees")
        case "wind":
            text=(f"The current wind condition in {city} is",wdic["current"]["wind_kph"])
        case "pressure":
            text=(f"The current pressure condition in {city} is", wdic["current"]["pressure_in"])
        case "humidity":
            text=(f"The current humidity in {city} is", wdic["current"]["humidity"])
        case "cloud":
            text=(f"The cloud condition in {city} is", wdic["current"]["cloud"])
        case "air quality":
            text=(f"The cloud condition in {city} is", wdic["current"]["air_quality"]["co"])
        case "local time":
            text=(f"The cloud condition in {city} is", wdic["location"]["localtime"])
        case _:
            text="wrong input"
            
    speak.Speak(text)        


    
    