import requests
import speech_recognition as s_r

from bs4 import BeautifulSoup
from chatterbot import ChatBot
from chatterbot.response_selection import get_random_response
from chatterbot.trainers import ChatterBotCorpusTrainer

import tkinter as tk
from tkinter import *
from tkinter.ttk import *

# Text-to-speech Libs
import win32com.client as wincl

speak = wincl.Dispatch("SAPI.SpVoice")
from sklearn import tree
import nltk
from gtts import gTTS
import os
# putting images in buttons
from tkinter import PhotoImage

try:
    import ttk as ttk
    import ScrolledText
except ImportError:
    import tkinter.ttk as ttk
    import tkinter.scrolledtext as ScrolledText
import time

"""
decision tree
"""
# solution to issue
solutions_data = ["Printer Test", "Router Test"]
# ["issue"]
issues = [["printer"], ["router"]]
# create decision classifier
decision_tree_classifier = tree.DecisionTreeClassifier()


class GUI(tk.Tk):

    def __init__(self, *args, **kwargs):
        """
        decision tree
        """
        # solution to issue
        self.solutions_data = ["Printer Test", "Router Test"]
        # ["issue"]
        self.issues = [["printer"], ["router"]]
        # create decision classifier
        self.decision_tree_classifier = tree.DecisionTreeClassifier()
        """
        Create & set window variables.
        """
        tk.Tk.__init__(self, *args, **kwargs)

        self.chatbot = ChatBot(
            "GUI Bot",
            response_selection_method=get_random_response,
            storage_adapter="chatterbot.storage.SQLStorageAdapter",
            logic_adapters=[
                "chatterbot.logic.BestMatch"
            ],
            database_uri="sqlite:///database.sqlite3"
        )
        # Corpus Trainer
        corpus = ChatterBotCorpusTrainer(self.chatbot)
        # insert corpus file
        corpus.train(
            # 'chatterbot.corpus.english.greetings',
            # 'chatterbot.corpus.english.issues'

        )

        # title
        self.title("Company Chatbot")
        # Gui
        self.initialize()

    # function for GUI
    def initialize(self):
        """
        ICON IMAGES
        """

        """
        Set window layout.
        """
        # specific format
        self.grid()
        # setting window size
        self.geometry('1200x1200')
        # color
        self.configure(bg='orange')
        # company chatbot logo
        aiIcon = PhotoImage(file="company_bot.png")
        self.co_bot = aiIcon.subsample(1, 1)
        self.co_bot_label = Label(self, image=self.co_bot)
        self.co_bot_label.grid(column=2, row=2)
        # speak
        # speak button
        speak = PhotoImage(file='chatbot_speak.png')
        self.speak = speak.subsample(1, 1)
        self.respond = ttk.Button(self, text='Chatbot Speak', command=self.ai_speech, image=self.speak)
        self.respond.grid(column=1, row=1, sticky='nesw', padx=3, pady=3)
        # create enter button for user
        # enter button
        enter = PhotoImage(file="enter_button.png")
        self.enter_icon = enter.subsample(2, 2)
        self.respond = ttk.Button(self, text='Press to Search Keyword:', command=self.get_response,
                                  image=self.enter_icon)
        self.respond.grid(column=0, row=0, sticky='nesw', padx=3, pady=3)
        # Speech recog
        record = PhotoImage(file="record.png")
        self.record_icon = record.subsample(1, 1)
        self.respond = ttk.Button(self, text='Speak To Chatbot', command=self.speech_recog, image=self.record_icon)
        self.respond.grid(column=0, row=1, sticky='nesw', padx=3, pady=3)
        # self.respond.config(width=70, height=70)
        # VPN
        self.respond = ttk.Button(self, text='How to connect to VPN', command=self.connect_vpn)
        self.respond.grid(column=0, row=6, sticky='nesw', padx=3, pady=3)
        # Monitor
        self.respond = ttk.Button(self, text='Monitor Issues', command=self.monitor_issues)
        self.respond.grid(column=1, row=3, sticky='nesw', padx=3, pady=3)
        # CoSoftware
        self.respond = ttk.Button(self, text='Software/Applications', command=self.co_software)
        self.respond.grid(column=1, row=4, sticky='nesw', padx=3, pady=3)
        # single_solutions
        self.respond = ttk.Button(self, text='Example: Single Solution', command=self.example_single)
        self.respond.grid(column=0, row=3, sticky='nesw', padx=3, pady=3)
        # multi_solutions
        self.respond = ttk.Button(self, text='Example: Multiple Solutions', command=self.example_multi)
        self.respond.grid(column=0, row=4, sticky='nesw', padx=3, pady=3)
        # step_solution
        self.respond = ttk.Button(self, text='Example: Step Solutions', command=self.example_step)
        self.respond.grid(column=0, row=5, sticky='nesw', padx=3, pady=3)
        # Passwords
        self.respond = ttk.Button(self, text='Password', command=self.passwords)
        self.respond.grid(column=0, row=7, sticky='nesw', padx=3, pady=3)
        # Printer
        self.respond = ttk.Button(self, text='Connecting to a Printer', command=self.local_printer_connect)
        self.respond.grid(column=1, row=5, sticky='nesw', padx=3, pady=3)
        # Syncing Issues
        self.respond = ttk.Button(self, text='Syncing Issues?', command=self.co_sync)
        self.respond.grid(column=1, row=6, sticky='nesw', padx=3, pady=3)
        # Trouble ticket
        self.respond = ttk.Button(self, text='Trouble Ticket', command=self.trouble_ticket)
        self.respond.grid(column=1, row=7, sticky='nesw', padx=3, pady=3)
        # keyword entry
        self.usr_input = ttk.Entry(self, state='normal')
        self.usr_input.grid(column=1, row=0, sticky='nesw', padx=3, pady=3)
        # conversation label for scrolledtext box

        # entry for user and bot conversation.
        self.conversation = ScrolledText.ScrolledText(self, width=100, wrap=WORD, state='disabled')
        self.conversation.grid(column=0, row=2, columnspan=2, sticky='nesw', padx=3, pady=3)
        # able scrolledtext to be edit
        self.conversation['state'] = 'normal'
        # insert directions on Chatbot
        self.conversation.insert(
            tk.END, " PLEASE READ!!"
                    "\n\n What is this: "
                    "\n  This a Company Chatbot is here to help you resolve your issues 24/7 without contacting IT Support!"
                    "\n\n What do I do:"
                    "\n  Whats your issue? See it on the screen? Click it and see what solutions are possible!"
                    "\n\n What are Keywords:"
                    "\n  Try not to type out the issue if you can't find a certain app type in Applications."
                    "\n  Try to use one or two words to resolve your issue Please!"
                    "\n\n No solution:"
                    "\n  Click Trouble Ticket and send a Trouble Ticket to the IT Support Team!")
        # disable scrolltext
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(" PLEASE READ!!"
                   "\n\n What is this: "
                   "\n  This a Company Chatbot is here to help you resolve your issues 24/7 without contacting IT Support!"
                   "\n\n What do I do:"
                   "\n  Whats your issue? See it on the screen? Click it and see what solutions are possible!"
                   "\n\n What are Keywords:"
                   "\n  Try not to type out the issue if you can't find a certain app type in Applications."
                   "\n  Try to use one or two words to resolve your issue Please!"
                   "\n\n No solution:"
                   "\n  Click Trouble Ticket and send a Trouble Ticket to the IT Support Team!")
        # closes file
        file.close()
        self.conversation['state'] = 'disabled'

    def ai_speech(self):
        # file input audio
        solution = open("speech.txt", "r").read().replace("\n", " ")
        # Playing the converted file
        speak.Speak(solution)

    def speech_recog(self):
        r = s_r.Recognizer()
        my_mic = s_r.Microphone(device_index=1)
        with my_mic as source:
            print("Say now!!!!")
            r.adjust_for_ambient_noise(source)  # reduce noise
            audio = r.listen(source)  # take voice input from the microphone
            # deletes last statements by bot bot and user
            self.delete_convo()
            # printed statement of voice input
            user_voice = (r.recognize_google(audio))

            self.usr_input.delete(0, tk.END)
            time.sleep(0.3)
            response = self.chatbot.get_response(user_voice)

            self.usr_input.insert(0, user_voice)
            self.conversation['state'] = 'normal'
            self.conversation.delete('1.0', tk.END)
            self.conversation.insert(
                tk.END, "You: " + user_voice + "\n" + "\n" + f"Bot: " + str(response) + "\n")
            self.conversation['state'] = 'disabled'
            # file is open and overwrites
            file = open("speech.txt", "w")
            # ai writes output in a file
            file.write(response.text)
            # closes file
            file.close()
            time.sleep(0.5)
            self.usr_input.delete(0, tk.END)

        print(r.recognize_google(audio))  # to print voice into te

    def trouble_ticket(self):
        """
       url = "https://google.org/crisisresponse/covid19-map"
        # url to be requested
        page = requests.get(url)
        # beautiful soup object is being placed(html.parser is important)
        soup = BeautifulSoup(page.content, 'html5lib')
        self.conversation.delete('1.0', tk.END)
        links = (soup.find_all('a'))
      """
        """
      Get a response from the chatbot and display it.
      """


    # delete entries in scrolledtext
    def delete_convo(self):
        # allows scrolltext box to be edit
        self.conversation['state'] = 'normal'
        # erases last input
        self.conversation.delete('1.0', tk.END)
        # scrolledtext can not be edit
        self.conversation['state'] = 'disabled'

    # bot response to user
    def get_response(self):
        """
        Get a response from the chatbot and display it.
        """
        # deletes last statements by bot bot and user
        self.delete_convo()
        # users input
        user_input = self.usr_input.get()
        # when user press enter
        # delete text
        self.usr_input.delete(0, tk.END)
        time.sleep(0.3)
        # output response according to user input
        response = self.chatbot.get_response(user_input)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "You: " + user_input + "\n" + "\n" + f"Bot: " + str(response) + "\n"
        )
        self.conversation['state'] = 'disabled'
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response.text)
        # closes file
        file.close()
        time.sleep(0.5)


    def example_single(self):
        self.delete_convo()
        resolve_issue1 = ["\n\n Solution: Easy fix try doing this"
                          ""
                          " \n\n Escalate: Trouble Ticket"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "" + str(response) + "\n")

    def example_multi(self):
        self.delete_convo()
        resolve_issue1 = ["\n\n Solution 1: How to resolve your issue"

                          "\n\n Solution 2: Follow what I say "

                          "\n\n Solution 3: Try different solutions"

                          "\n\n Solution 4: If previous solution does not work"

                          "\n\n Solution 5: Escalate to Trouble Ticket"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "" + str(response) + "\n")

    def example_step(self):
        self.delete_convo()
        resolve_issue1 = [
            "\n\n Step 1: How to resolve your issue      "

            "\n\n Step 2: Follow what bot has for you     "

            "\n\n Step 3: Try different solution     "

            "\n\n Step 4: If previous solution does not work      "

            "\n\n Step 5: Create Trouble Ticket"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "" + str(response) + "\n")

    def passwords(self):
        self.delete_convo()
        resolve_issue1 = ["\n\nPassword Solutions:"

                          "\n\n Solution 1: Default password is tech001"

                          "\n\n Solution 2: Try using your tech number instead of 001"

                          "\n\n Solution 3: If your a new hire give 48-72 hours for your account to be process"

                          "\n\n Escalate: Trouble Ticket"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "How to resolve your issue try: " + str(response) + "\n")

    def co_sync(self):
        self.delete_convo()
        resolve_issue1 = ["\n\nCompnay Sync Steps: "
                          "\n\n Step 1: Locate File Explorer in your Taskbar (Folder Icon) "
                          "\n\n Step 2: Once in File Explorer, Find ' This PC' "
                          "\n\n Step 3: Under 'This PC' click on 'Windows(C:)"
                          "\n\n Step 4: Click on Application Folder"
                          "\n\n Step 5: Click on Company Folder"
                          "\n\n Step 6: There are two folders, delete the data folder"
                          "\n\n Step 7: Relaunch Company Sync"
                          "\n\n Escalate: Trouble Ticket"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "How to resolve your issue try: " + str(response) + "\n")

    def local_printer_connect(self):
        self.delete_convo()
        resolve_issue1 = ["\n\nConnecting to a Local Printer: "
                          "\n\n Step 1: In your Taskbar search: 'Control Panel'  "
                          "\n\n Step 2: Double-click Printers"
                          "\n\n Step 3: Double-click Add printer icon "
                          "\n\n Step 4: Click next to start the Add a printer wizard"
                          "\n\n Step 5: Select Network Printer, click Next"
                          "\n\n Step 6: Select a shared printer by name: 'companyprinter123'"
                          "\n\n Escalate: Trouble Ticket"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "How to resolve your issue try: " + str(response) + "\n")

    def connect_vpn(self):
        self.delete_convo()
        resolve_issue1 = ["\n\nConnecting to VPN: \n\n Step 1: Search Global Protect in the search bar on your Taskbar "
                          "\n\n Step 2: Double Click Global Protect App to run it "
                          "\n\n Step 3: Enter you Work Email and Password"
                          "\n\n POSSIBLE ISSUES WITH SOLUTIONS"
                          "\n\n Issue1: Can't Find your Taskbar?"
                          "\n Solution: Taskbar is where you find icons to apps, date, and time.  "
                          "\n\n Issue2: Can't Find Global Protect App?"
                          "\n Solution: Step 1: In your search bar type software applications Click Software Applications"
                          "\n           Step 2: Locate Global Protect Applcation and double click to download"
                          "\n           Step 3: Once downloaded it will open. "
                          "\n           Step 4: Enter your Work Email and Password"]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "" + str(response) + "\n")

    def monitor_issues(self):
        self.delete_convo()
        resolve_issue1 = ["\n\nHaving Monitor Issues: "
                          "\n\n Solution 1: Make sure cords in back of monitor are tight."
                          "\n\n Solution 2: Double Check all cords are tighten! "
                          "\n\n Solution 3: If one monitor is working try logging in if you haven't logged in yet"
                          "\n\n POSSIBLE ISSUES WITH SOLUTIONS"
                          "\n\n Issue1: Logged in but only one monitor is working."
                          "\n Solution: Follow these STEPS:  "
                          "\n                        Step 1: In your taskbar in the Windows Search type: 'Display' "
                          "\n                        Step 2: Click on 'Display Settings' "
                          "\n                        Step 3: You should see Numbers 1 and 2 Scroll down until you see: 'Multiple "
                          "                          displays'"
                          "\n                        Step 4:  In the drop down click 'Extend these displays' "
                          "\n\n Escalate: Could be bad wires submit a Trouble Ticket"
                          ]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "" + str(response) + "\n")

    def co_software(self):
        self.delete_convo()
        resolve_issue1 = ["\n\nCompany Software: "
                          "\n\n You have access to download certain software: "
                          "\n\n In your Taskbar search: 'Software Applications'"
                          "\n\n Once your in Software Applications you will see all the software available to download "
                          "\n with no extra credentials "
                          "\n\n Escalate: If you don't see the software needed. Put in a Trouble Ticket to request the software"
                          "\n you would like to use to see if it is allowed to be downloaded! "
                          ]
        # Create response
        response = "".join(resolve_issue1)
        # file is open and overwrites
        file = open("speech.txt", "w")
        # ai writes output in a file
        file.write(response)
        # closes file
        file.close()
        # delete inputed data
        self.usr_input.delete(0, tk.END)
        self.conversation['state'] = 'normal'
        self.conversation.insert(
            tk.END, "" + str(response) + "\n")


gui = GUI()
gui.mainloop()
