#find junk emails from inbox and move them to Junk mail folder

import win32com.client
import re
import yaml

stream = open("settings.yaml", "r")
params = yaml.load(stream)

_accountName = params["outlook_account"]
_inbox = params["outlook_inbox"]
_junk = params["outlook_junk"]

text_file = open(params["keywords_file"], "r", encoding="utf-8")
_junkKeyWords = text_file.readlines()

moved = 0
items = []

def Contains(searchStr, inStr):
    searchStr = searchStr.strip()
    if searchStr == '': #empty line always return false
        return False
    else:
        return re.search("(?i)"+searchStr, inStr, re.IGNORECASE)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders(_accountName).Folders(_inbox)
junk = outlook.Folders(_accountName).Folders(_junk)

messages = inbox.Items

for message in list(messages):
    for junkKeyWord in _junkKeyWords:
        if Contains(junkKeyWord, message.Subject) or Contains(junkKeyWord, message.Body):
            message.Move(junk)
            moved += 1
            break
print()
print("Found and moved " + str(moved) + " junk emails")

