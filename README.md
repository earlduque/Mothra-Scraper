# Mothra-Scraper
Publish Date: 9/6/2017
Updated: 9/13/2017
Created by: eduque

All included files are open source, the VBA code in the excel file is not locked.

Download all files in this repo.

Open Mothra Screens.xlsm

	-Enable Macros by selecting "Enable Content"
	-Say "Yes" to the security warning
	-This file is where we'll place all the information scraped from Mothra
	
Open Putty and ssh into Mothra

	-It can be any window size and titled whatever you want

Run one of the Scraper files:

	-If you have AutoHotKey installed, you can run Mothra Scraper.ahk
	-If you don't have AutoHotKey installed, you can run the .exe

	1. When you run either of these files, it will prompt you for your Mothra window title
	2. Enter the title EXACTLY as it appears on the top of your Putty Mothra Window, and then continue
	3. A confirmation will appear and you will be shown what Hotkeys are now available to you and what they do
	
In order to run the scraper:

	-Get to the Mothra landing page (open putty, bastion in, log in)
	-Go to Display Functions (D)
	-Go to Display MailID (F)
	-Type in a MailID and press enter, a valid user record should appear
	-Press Windows+Shift+M
	
This app also adds the following functionality:

	-Control+C no longer closes Mothra
	-Control+V actually pastes
	-Contral+A selects all
