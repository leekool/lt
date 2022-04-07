scripts i've created to transcribe court hearings and to learn python



### lt.py

lt.py is a CLI with the following commands:

initials - change the value of 'initials' in config.json (the user's initials) (requires argument 'initials')

prefix - changes the value of 'prefix' in config.json (the beginning of the name of a turn document before the turn number) (requires argument 'prefix')

s1, s2, s3, s4 - change the values of speakers 1, 2, 3, and 4 in config.json accordingly

doc - requires argument 'turn', creates a new document with the turn template corresponding with 'turn' specified, writes its name and path as 'last_turn' and 'last_turn_path' to config.json.

daily - prints list of folders in the current date's folder and allows you to choose one, opening the folders and changing 'daily_path' in config.json accordingly, attempts to find the running sheet in the folder - converts .doc running sheets to .docx -  and pull 'prefix' from it and print turns corresponding with 'initials'

vpn - toggles connection to the legal transcripts VPN

save - saves and closes 'last_turn', counts the words, writes 'last_turn', today's date, and the word count into the next empty row of the excel invoice, then moves the turn to 'daily_path'



requires: 
- click
- docx
- pywinauto
- openpyxl


### lt.ahk

lt.ahk is a autohotkey script with keybindings that send speaker names, questioning and answering, etc

uses [AutoHotkey-JSON](https://github.com/cocobelgica/AutoHotkey-JSON) to read config.json


