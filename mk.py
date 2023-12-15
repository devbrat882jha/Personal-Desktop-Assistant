import speech_recognition as sr 
import os
import win32com.client 
import webbrowser
import requests
import time
import datetime
import random

def speak(s):
      speaker = win32com.client.Dispatch("SAPI.SpVoice") 
      x=speaker.Speak(s)
      return x
speak("hello i am your personal assistant")

def speech_conversion():
  recogniser=sr.Recognizer()#create instance of recogniser class which is core library
  try:
    with sr.Microphone() as mic:#opens the micrphone,with statment helps to close mic when block os exit thats why with is important,thus in this stage voice is heard
      audio=recogniser.listen(mic,timeout=2)#listen function returns audio data
      text=recogniser.recognize_google(audio,language='en-in')#converts audio data into string text with help of google
      print(f"user said:{text}")
      return text
  except Exception as e:
     return "can't recognise voice"
  

def timestamp(x):#to convert weather data into date time in srting format
    dt_object= datetime.datetime.fromtimestamp(x)
    string_time = dt_object.strftime('%Hhours:%Mminutes:%Sseconds')
    return string_time



while True:   
   text2=speech_conversion()

   if 'exit' in text2.lower():
     exit()

   if 'create folder' in text2.lower():
       
     speak("what name do you like to folder")
     
     new_text=speech_conversion().lower().replace(' ', '_')
     if new_text=="can't_recognise_voice":
        print("say again")
        new_text=speech_conversion().lower().replace(' ','_')
     if new_text!="can't_recognise_voice":
      if not os.path.exists(f"{new_text}"):
        os.mkdir(f"{new_text}")
   if 'create file' in text2.lower():
     speak('what name do you want of file  ')
     response=speech_conversion().lower().replace(' ','_')
     print(response)
     if response=="can't_recognise_voice":
        print("say again")
        response=speech_conversion().lower().replace(' ','_')
     if response!="can't_recognise_voice":
       open(f"{response}",'w').close()

   if 'open site' in text2.lower():
     speak('what site you wish to open')
     response2=speech_conversion().lower().replace(' ','_')
     print(response2)
     if response2=="can't_recognise_voice":
        print("say again")
        response2=speech_conversion().lower().replace(' ','_')
     if response2!="can't_recognise_voice":
       webbrowser.open(f"https://www.{response2}.com")
   if 'open movie' in text2.lower():
     speak('which movie')
     response3=speech_conversion().lower().replace(' ','_')
     print(response3)
     if response3=="can't_recognise_voice":
        print("say again")
        response3=speech_conversion().lower().replace(' ','_')
     if response3!="can't_recognise_voice":
       os.startfile(f"C:\\Users\\devbr\\Downloads\\movies\\{response3}.mp4")

   if 'headlines'  in text2.lower():
      

      api_key='api key'
      url=f'https://newsapi.org/v2/top-headlines?sources=techcrunch&apiKey={api_key} '


      req=requests.get(url).json()
      
      art=req['articles']
      new_art=[]
      for i in art:
          new_art.append(i['title'])
      for j in new_art:
        speak(j)
   if 'the time' in text2.lower():
      t=time.strftime("%Hhours%Mminutes%Sseconds")
      speak(t)
   if 'weather' in text2.lower():
      api='apik key'
      lat='19.07'
      lon='72.87'
      url1=f'https://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={api}'
      req1=requests.get(url1).json()
     
      temp=req1['main']['temp']-273.15
      temp=round(temp,1)
      tempmin=req1['main']['temp_min']-273.15
      tempmin=round(tempmin,1)
      tempmax=req1['main']['temp_max']-273.15
      tempmax=round(tempmax,1)
      pressure=req1['main']['pressure']
      humidity=req1['main']['humidity']
      sunr=req1['sys']['sunrise']
      sunrise=timestamp(sunr)
      suns=req1['sys']['sunset']
      sunset=timestamp(suns)
      visibility=req1['visibility']

      values_of_elements=[temp,tempmin,tempmax,pressure,humidity,sunrise,sunset,visibility]
      weather_elements=['temperature','temperature minimum','temperature max','pressure','humidity','sunrise','sunset','visibility']
      for name,measurement in zip(weather_elements,values_of_elements):
         speak(f"{name} is {measurement}")


   if 'write for me' in text2.lower():
      speak('okay tell me sir what i shall write for you ')
      response4=speech_conversion().lower().replace(' ','_')
      if response4=="can't_recognise_voice":
        response4=speech_conversion().lower().replace(' ','_')
      if response4!="can't_recognise_voice":
         response4=response4.replace('_',' ')
         alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
         random_letter = random.choice(alphabet)
         
         with open(f'{random_letter}.txt','w') as d:
            d.write(f'{response4}')
            d.close()
      
