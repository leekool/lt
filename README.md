scripts i've created to transcribe court hearings and to learn python

---

lt.py is a CLI with the following commands:

pf - changes the value of 'prefix' in config.json (the beginning of the name of a turn document before the turn number)

s1, s2, s3, s4 - change the values of speakers 1, 2, 3, and 4 in config.json accordingly

doc - requires argument 'turn', creates a new document with the turn template corresponding with 'turn' specified, writes its name and path as 'last_turn' and 'last_turn_path' to config.json

daily - shows a list of folders in the current date's folder and allows you to choose one, opening the folders and changing 'daily_path' in config.json accordingly

vpn - toggles connection to the legal transcripts VPN

save - saves and closes 'last_turn', counts the words, writes 'last_turn', today's date, and the word count into the next empty row of the excel invoice, then moves the turn to 'daily_path'

---

lt.ahk is a autohotkey script for keybindings that sends speaker names, questioning and answering, etc

---

want to add means to read a daily folder's running sheet .docx and set 'pf' automatically 
need to create a variable and command for changing the excel invoice name and path
