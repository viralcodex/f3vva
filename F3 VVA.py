import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import urllib.parse

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

uname=""

camera_port=4
def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !")

	else:
		speak("Good Evening Sir !")

	assname =("Jarvis 1 point o")
	speak("I am your Assistant")
	speak(assname)


def usrname():
	speak("What should i call you sir")
	uname = takeCommand()
	speak("Welcome Mister")
	speak(uname)
	columns = shutil.get_terminal_size().columns

	print("###############################################################".center(columns))
	print("Welcome Mr.", uname.center(columns))
	print("###############################################################".center(columns))

	speak("How can i Help you,"+uname)

def takeCommand():

	r = sr.Recognizer()

	with sr.Microphone() as source:

		print("Listening...")
		r.pause_threshold = 1
		audio = r.listen(source)

	try:
		print("Recognizing...")
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e)
		print("Unable to Recognizing your voice.")
		return "None"

	return query

'''def sendEmail(to, content):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()

	# Enable low security in gmail
	server.login('your email id', 'your email passowrd')
	server.sendmail('your email id', to, content)
	server.close()
'''
if __name__ == '__main__':
	clear = lambda: os.system('cls')

	# This Function will clean any
	# command before execution of this python file
	clear()
	wishMe()
	usrname()

	while True:

		query = takeCommand().lower()

		# All the commands said by user will be
		# stored here in 'query' and will be
		# converted to lower case for easily
		# recognition of command
		if 'wikipedia' in query:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences = 3)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		elif 'open youtube' in query:
			speak("Here you go to Youtube\n")
			webbrowser.open("https://www.youtube.com")

		elif 'open google' in query:
			speak("Here you go to Google\n")
			webbrowser.open("https://www.google.com")

		elif 'open stackoverflow' in query:
			speak("Here you go to Stack Over flow.Happy coding")
			webbrowser.open("https://www.stackoverflow.com")

		elif 'play music' in query or "play song" in query:
			flag=0
			speak("Here you go with music")
			music_dir ="C:\\Users\\amanm\\Desktop\\project\\music"
			#music_dir = "C:\\Users\\HP\\Music"
			songs = os.listdir(music_dir)
			print(songs)
			speak("choose a song")
			query= takeCommand().lower()
			for s in songs:
				if s.lower()==query+".mp3":
					random = os.startfile(os.path.join(music_dir, s))
					flag=1
			if flag==0:
				speak("song not found")

		elif 'time' in query:
			strTime = datetime.datetime.now().strftime("%H:%M:%S")
			print(strTime)
			speak(f"Sir, the time is {strTime}")

		elif 'open song' in query:
			codePath = r"C:\\Users\\amanm\\Desktop\\project\\music\\Perfect.mp3"
			os.startfile(codePath)

		elif 'email to gaurav' in query:
			try:
				speak("What should I say?")
				content = takeCommand()
				to = "Receiver email address"
				sendEmail(to, content)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		elif 'send a mail' in query:
			try:
				speak("What should I say?")
				content = takeCommand()
				speak("whome should i send")
				to = input()
				sendEmail(to, content)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you,"+uname)

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			assname = query

		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query or "What is your name" in query:
			speak("My friends call me")
			speak(assname)
			print("My friends call me", assname)

		elif 'goodbye' in query or 'see you later' in query or 'bye' in query or 'exit' in query:
			speak("Thanks for giving me your time")
			exit()

		elif "who made you" in query or "who created you" in query:
			speak("I have been created by Team F3.")

		elif 'joke' in query:
			speak(pyjokes.get_joke())

		elif "calculate" in query:

			app_id = "485437-UKK6HRJHPJ"
			client = wolframalpha.Client(app_id)
			indx = query.lower().split().index('calculate')
			query = query.split()[indx + 1:]
			res = client.query(query)
			try:
				answer = next(res.results).text
			except StopIteration:
				answer="no result"
			print("The answer is " + answer)
			speak("The answer is " + answer)

		elif 'search' in query or 'play' in query:

			query = query.replace("search", "")
			query = query.replace("play", "")
			base_url = "http://www.google.com/?#q="
			final_url = base_url + urllib.parse.quote(query)
			webbrowser.open(final_url )

		elif "who i am" in query:
			speak("If you talk then definitely you're human.")

		elif 'power point presentation' in query:
			speak("opening Power Point presentation")
			power = r"C:\\Users\\HP\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
			os.startfile(power)

		elif 'is love' in query:
			speak("It is 7th sense that destroy all other senses")

		elif "who are you" in query:
			speak("I am your virtual assistant created by Team F3")

		elif 'reason for you' in query or "why you came to world" in query:
			speak("I was created as a Minor project by Team F3 ")

		elif 'change background' in query:
			ctypes.windll.user32.SystemParametersInfoW(20,
													0,
													"Location of wallpaper",
													0)
			speak("Background changed succesfully")

		elif 'open python' in query:
			appli = r"C:\\py\\python.exe"
			os.startfile(appli)

		elif 'news' in query:

			try:
				jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
				data = json.load(jsonObj)
				i = 1

				speak('here are some top news from the times of india')
				print('''=============== TIMES OF INDIA ============'''+ '\n')

				for item in data['articles']:

					print(str(i) + '. ' + item['title'] + '\n')
					print(item['description'] + '\n')
					speak(str(i) + '. ' + item['title'] + '\n')
					i += 1
			except Exception as e:

				print(str(e))


		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')

		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		elif "don't listen" in query or "stop listening" in query:
			speak("for how much time you want to stop jarvis from listening commands")
			a = takeCommand()
			time_list=a.split(" ")
			if time_list[1].lower()=="seconds" or time_list[1].lower()=="second":
				pass
			elif time_list[1].lower()=="minutes" or time_list[1].lower()=="minute":
				time_list[0]=int(time_list[0])*60
			elif time_list[1].lower() == "hours" or time_list[1].lower() == "hour":
				time_list[0] = int(time_list[0]) * 3600
			print("meet you in "+time_list[:2])
			speak("meet you in "+time_list[:2])
			time.sleep(int(time_list[0]))

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/" + location + "")


		elif "camera" in query or "take a photo" in query:
			ec.capture(0, "Jarvis Camera ", "img.jpg")
			
		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])

		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "write a note" in query:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('jarvis.txt', 'w')
			speak("Sir, Should i include date and time")
			snfm = takeCommand()
			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime("% H:% M:% S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
			else:
				file.write(note)

		elif "show note" in query:
			speak("Showing Notes")
			file = open("jarvis.txt", "r")
			print(file.read())
			speak(file.read(6))

		elif "update assistant" in query:
			speak("After downloading file please replace this file with the downloaded one")
			url = '# url after uploading file'
			r = requests.get(url, stream = True)

			with open("Voice.py", "wb") as Pypdf:

				total_length = int(r.headers.get('content-length'))

				for ch in progress.bar(r.iter_content(chunk_size = 2391975),
									expected_size =(total_length / 1024) + 1):
					if ch:
						Pypdf.write(ch)

		# NPPR9-FWDCX-D2C8J-H872K-2YT43
		elif "jarvis" in query:

			wishMe()
			speak("Jarvis 1 point o in your service Mister")
			speak(assname)

		elif "weather" in query:            # Google Open weather website to get API of Open weather
			speak(" City name ")
			print("City name : ")
			CITY = takeCommand()
			# base URL
			BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
			API_KEY = "1e17359ab6ad6b1fb9794d77fcfb5dc2"
			# upadting the URL
			URL = BASE_URL + "q=" + CITY + "&appid=" + API_KEY
			# HTTP request
			response = requests.get(URL)
			# checking the status code of the request
			if response.status_code == 200:
				# getting data in the json format
				data = response.json()
				# getting the main dict block
				main = data['main']
				# getting temperature
				temperature = main['temp']
				# getting the humidity
				humidity = main['humidity']
				# getting the pressure
				pressure = main['pressure']
				# weather report
				report = data['weather']
				print(f"{CITY:-^30}")
				print(f"Temperature: {temperature}")
				print(f"Humidity: {humidity}")
				print(f"Pressure: {pressure}")
				print(f"Weather Report: {report[0]['description']}")
			else:
				speak(" City Not Found ")

		elif "send message " in query:
				# You need to create an account on Twilio to use this service
				account_sid = 'Account Sid key'
				auth_token = 'Auth token'
				client = Client(account_sid, auth_token)

				message = client.messages.create(
									body = takeCommand(),
									from_='Sender No',
									to ='Receiver No'
								)

				print(message.sid)

		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Good Morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assname)

		# most asked question from google Assistant
		elif "will you be my gf" in query or "will you be my bf" in query:
			speak("I'm not sure about, may be you should give me some time")

		elif "how are you" in query:
			speak("I'm fine, glad you me that")

		elif "i love you" in query:
			speak("It's hard to understand")

		elif "what is" in query or "who is" in query:

			# Use the same API key
			# that we have generated earlier
			client = wolframalpha.Client("API_ID")
			res = client.query(query)

			try:
				print (next(res.results).text)
				speak (next(res.results).text)
			except StopIteration:
				print ("No results")

		# elif "" in query:
			# Command go here
			# For adding more commands



