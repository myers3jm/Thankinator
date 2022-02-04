# This is a Python script written in 2019 for the purpose of mass-writing thank-you notes.
# The program was written for the occasion of my high school graduation, and it reflects
# some of the experience I had with Python at the time. I documented the notes, gifts, and
# other kindnesses I received at my graduation party in an Excel spreadsheet and instructed
# the program to pull information from that spreadsheet. The program took into account
# several variables, including whether or not a person sent a gift (and what that gift was),
# whether or not a person was at my graduation party, whether or not they wrote a note in a
# book my mother put together, and whether or not they sent a card. Based on these variables,
# the program wrote a unique letter of thanks to each person in the spreadsheet. I did not
# time the process as it ran with much accuracy, but it was done writing and printing
# approximately 150 letters in 10 minutes.

# Since the creation of this script, I have seen it evolve through my mother's use of its
# underlying functionalities to send notes to the people in her life on occasions such as
# birthdays (even receiving one myself).

# Looking back on the script, I believe that if I were to face the same task again, I
# would take an almost entirely different approach. I would likely write the sequel in Python
# as well, due to the language's simplicity of use, my own familiarity with its syntax, and
# the plethora of libraries that the community has put forth for various tasks. I would
# change the means through which the program interfaced with the output letters, likely
# opting to write to a text or Word file directly through the use of the filestream or a
# library specifically created for Word documents. The script below instead sends keystrokes
# that correspond to the specific shortcuts to switch between Excel and Google Chrome on my
# computer, and then sending the keystrokes that match the contents of the letter to a
# Google Doc. At the time I thought this approach was very sophisticated, though now I guess
# that the total execution time of the process could be lowered from 150 letters in 10 minutes
# to 150 letters in 3-5, depending on the speed of the connected printer.



##### ALL CODE BELOW THIS LINE IS THE ORIGINAL PROGRAM #####



#Import and establish workbook/worksheet
import xlrd
import pyautogui as pygui
import time
workbook = xlrd.open_workbook("C:\\Users\\Jared\\Downloads\\GradPartyLoot.xlsx")
sheet = workbook.sheet_by_name("Sheet1")

#User decision of range or list
range_list = input("Thankinate on a range (R) or from a list (L): ")

#Customization break
def cust():
	input("press ENTER when ready to print")
	pygui.hotkey('win', '2')

#Range
if str(range_list.lower()) == "r":
	#Loop through column 0, retrieving data until the given cell
	start = input("Start Value: ")
	end = input("End Value: ")

	for x in range(int(start) - 1, int(end)):
		if sheet.cell(x,0).value != xlrd.empty_cell.value:
			person = str(sheet.cell(x,0).value)
			gift = str(sheet.cell(x,2).value)
			if sheet.cell(x,3).value != xlrd.empty_cell.value:
				note = True
			if sheet.cell(x,3).value == xlrd.empty_cell.value:
				note = False
			if sheet.cell(x,4).value != xlrd.empty_cell.value:
				mail = True
			if sheet.cell(x,4).value == xlrd.empty_cell.value:
				mail = False
			#Generating the note
			letter = str()
			letter += ("Dear " + person.title() + ",\n")
			#If they mailed or not
			if mail:
				letter += ("	Thank you so much for congratulating me! I would not be who I am today were it not for the friends, family, and teachers I've had the privilege to know. I am so grateful for your role in my life. ")
			if not mail:
				letter += ("	Thank you so much for coming to my graduation party! I would not be who I am today were it not for the friends, family, and teachers I've had the privilege to know. I am so grateful for your role in my life. ")
			#If they wrote a note or notes
			if len(sheet.cell(x,3).value) == 1:
				letter += ("Additionally, I want to thank you for your note in \"Oh the Places You'll Go!\" Seeing all the kind things everyone said really made graduating feel that much more special. ")
			if len(sheet.cell(x,3).value) > 1:
				letter += ("Additionally, I want to thank you for your notes in \"Oh the Places You'll Go!\" Seeing all the kind things everyone said really made graduating feel that much more special. ")
			#If they gave a card or gift
			if len(str(sheet.cell(x,2).value)) > 0:
				if str(gift) != "card":
					if gift[-1] == "0":
						letter += ("I'd also like to thank you for your generosity, I will be sure to put the $" + str(gift) + "0 to good use! ")
					else:
						letter += ("I'd also like to thank you for your generosity, I will be sure to put the $" + str(gift) + " to good use! ")
				elif str(gift) == "card":
					letter += ("I'd also like to thank you for the card you gave me! ")
			
			letter += ("Most importantly, I want to thank you for your support as I've grown up. There's just no telling who I'd be if it weren't for you.\n")
			letter += ("\n		Sincerely,")
			#Typing
			
			print("Row " + str(x + 1) + ": " + str(x + 1) + "/" + str(end) + " " + str(person))
			pygui.hotkey('win', '2')
			pygui.hotkey('ctrl', 'a')
			pygui.hotkey('backspace')
			pygui.hotkey('enter')
			pygui.typewrite(str(letter))
			print(str(letter) + "\n")
			pygui.hotkey('enter')
			pygui.hotkey('enter')
			cust()
			time.sleep(3)
			pygui.hotkey('ctrl', 'p')
			time.sleep(2)
			pygui.hotkey('enter')
			pygui.hotkey('win', '1')

	input("All Done!")


#List
if str(range_list.lower()) == "l":
	list = []
	print("Enter table values to add them to the list. Once finished, type \"done\"")
	cont = True
	while cont:
		addition = input("Value: ")
		if str(addition).lower() != "done":
			list.append(int(addition) - 1)
		else:
			cont = False
	
	i = 1
	#Loop through column 0, retrieving data until the given cell
	for x in list:
		if sheet.cell(x,0).value != xlrd.empty_cell.value:
			person = str(sheet.cell(x,0).value)
			gift = str(sheet.cell(x,2).value)
			if sheet.cell(x,3).value != xlrd.empty_cell.value:
				note = True
			if sheet.cell(x,3).value == xlrd.empty_cell.value:
				note = False
			if sheet.cell(x,4).value != xlrd.empty_cell.value:
				mail = True
			if sheet.cell(x,4).value == xlrd.empty_cell.value:
				mail = False
			#Generating the note
			letter = str()
			letter += ("Dear " + person.title() + ",\n")
			#If they mailed or not
			if mail:
				letter += ("	Thank you so much for congratulating me! I would not be who I am today were it not for the friends, family, and teachers I've had the privilege to know. I am so grateful for your role in my life. ")
			if not mail:
				letter += ("	Thank you so much for coming to my graduation party! I would not be who I am today were it not for the friends, family, and teachers I've had the privilege to know. I am so grateful for your role in my life. ")
			#If they wrote a note or notes
			if len(sheet.cell(x,3).value) == 1:
				letter += ("Additionally, I want to thank you for your note in \"Oh the Places You'll Go!\" Seeing all the kind things everyone said really made graduating feel that much more special. ")
			if len(sheet.cell(x,3).value) > 1:
				letter += ("Additionally, I want to thank you for your notes in \"Oh the Places You'll Go!\" Seeing all the kind things everyone said really made graduating feel that much more special. ")
			#If they gave a card or gift
			if len(str(sheet.cell(x,2).value)) > 0:
				if str(gift) != "card":
					if gift[-1] == "0":
						letter += ("I'd also like to thank you for your generosity, I will be sure to put the $" + str(gift) + "0 to good use! ")
					else:
						letter += ("I'd also like to thank you for your generosity, I will be sure to put the $" + str(gift) + " to good use! ")
				elif str(gift) == "card":
					letter += ("I'd also like to thank you for the card you gave me! ")
			
			letter += ("Most importantly, I want to thank you for your support as I've grown up. There's just no telling who I'd be if it weren't for you.\n")
			letter += ("\n		Sincerely,")
			#Typing
			
			print("Row " + str(x + 1) + ": " + str(i) + "/" + str(len(list)) + " " + str(person))
			pygui.hotkey('win', '2')
			pygui.hotkey('ctrl', 'a')
			pygui.hotkey('backspace')
			pygui.hotkey('enter')
			pygui.typewrite(str(letter))
			print(str(letter) + "\n")
			pygui.hotkey('enter')
			pygui.hotkey('enter')
			cust()
			time.sleep(3)
			pygui.hotkey('ctrl', 'p')
			time.sleep(2)
			pygui.hotkey('enter')
			pygui.hotkey('win', '1')
			i += 1

	input("All Done!")