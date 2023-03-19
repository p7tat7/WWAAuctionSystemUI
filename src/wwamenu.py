# -*- coding: UTF-8 -*-

import wwaglobal
import wwalog
import wwautils
import wwafindbidder
import tkinter as tk
wwaglobal.init()
def showMenu(options):
	for i in range(len(options)):
		print(str(i) + ". " + options[i])
	while 1:
		temp = input("> ")
		try:
			choice = int(temp)
			if choice < 0 or choice > len(options) - 1:
				print("Invalid input.")
			else:
				print()
				return choice
		except:
			print("Invalid input.")

def mainMenu():
	pass
	# while 1:
	# 	choice = showMenu(["Exit", "拍賣系統", "CMS(網店)","Whatsapp功能","催拍賣中標者下單","拍賣品位置","test"])
	#
	# 	if choice == 1:
	# 		# Auction
	# 		auctionMenu()
	# 	elif choice == 2:
	# 		# CMS
	# 		cmsMenu()
	# 	elif choice == 3:
	# 		whatsappFunction()
	# 	elif choice == 4:
	# 		wwautils.urgeOrder()
	# 	elif choice == 5:
	# 		wwautils.check_auction_item()
	# 	elif choice == 6:
	# 		wwautils.test()
	# 	else:
	# 		# Exit
	# 		break
	# print("Bye!")

	#Windows Creation

# class main_menu(tk.Tk):
# 	def __init__(self, *args, **kwargs):
# 		tk.Tk.__init__(self, *args, **kwargs)
# 		container = tk.Frame(self)
# 		container.pack(side='top', fill='both', expand = True)
# 		container.grid_rowconfigure(0, weight=1)
# 		container.grid_columnconfigure(0, weight=1)
# 		self.frames = {}
# 		for F in (menu, auction_menu):
# 			page_name = F.__name__
# 			frame = F(parent=container, controller=self)
# 			self.frames[page_name] = frame
# 			frame.grid(row=0,column=0,sticky='nsew')
# 		self.show_frame('menu')
# 	def show_frame(self, page_name):
# 		frame = self.frames[page_name]
# 		frame.tkraise()
#
# class menu(tk.Frame):
# 	def __init__(self, parent, controller):
# 		tk.Frame.__init__(self, parent)
# 		self.controller = controller
# 		# Windows Creation
# 		window = tk.Tk()
# 		window.title('Auction System')
# 		window.geometry('500x500')
# 		lbl_1 = tk.Label(self, text='Auction System')
# 		lbl_1.grid(column=0, row=0)
#
# 		# Auction Menu Button
# 		btn_1 = tk.Button(self, text='拍賣系統')
# 		btn_1['width'] = 50
# 		btn_1['height'] = 4
# 		btn_1.grid(column=0, row=2)
# 		btn_1['command'] = controller.show_frame('auction_menu')
#
# 		# CMS Menu
# 		btn_2 = tk.Button(self, text='CMS 網店操作')
# 		btn_2['width'] = 50
# 		btn_2['height'] = 4
# 		btn_2.grid(column=0, row=3)
# 		btn_2['command'] = cmsMenu
#
# 		# Whatsapp Menu
# 		btn_3 = tk.Button(self, text='Whatsapp 功能')
# 		btn_3['width'] = 50
# 		btn_3['height'] = 4
# 		btn_3.grid(column=0, row=2)
# 		btn_3['command'] = whatsappFunction
#
#
# class auction_menu(tk.Frame):
# 	def __init__(self, parent, controller):
# 		tk.Frame.__init__(self, parent)
# 		self.controller = controller
# 		label = tk.Label(self, text='拍賣系統')
# 		label.pack(side='top', fill='x', paddy=10)


def auctionMenu():
	def windes():
		window_auction.destroy()
	window_auction = tk.Toplevel(window)
	window_auction.geometry('500x500')
	window_auction.title('拍賣')

	# Back Button
	btn_1 = tk.Button(window_auction, text='返回')
	btn_1['width'] = 50
	btn_1['height'] = 4
	btn_1.grid(column=0, row=1)
	btn_1['command'] = windes

	# Post Auction Menu
	btn_2 = tk.Button(window_auction, text='出拍賣')
	btn_2['width'] = 50
	btn_2['height'] = 4
	btn_2.grid(column=0, row=2)
	btn_2['command'] = wwautils.postAuction

	# Whatsapp Menu
	btn_3 = tk.Button(window_auction, text='Whatsapp 功能')
	btn_3['width'] = 50
	btn_3['height'] = 4
	btn_3.grid(column=0, row=3)
	btn_3['command'] = whatsappFunction
	# while 1:
	# 	#choice = showMenu(["Back", "Post auction", "Share auction", "End auction"])
	# 	choice = showMenu(["Back", "出拍賣", "確認中標者 請先確認中標者再完標", "完標"])
	#
	#
	# 	if choice == 1:
	# 		# post auction
	# 		wwalog.log("--Post Auction Start--")
	# 		wwautils.postAuction()
	# 		wwalog.log("--Post Auction End--")
	# 	elif choice == 2:
	# 		# Find bidder and price (before end auction)
	# 		wwalog.log("--Find bidder and price (before end auction) Start--")
	# 		nameList = wwafindbidder.findbidder()
	# 		f = open("nameList.txt", "w+", encoding = "utf-8")
	# 		for i in range(len(nameList)):
	# 			if not nameList[i] == "":
	# 				f.write(nameList[i] + "\n")
	# 		f.close()
	# 		wwalog.log("--Find bidder and price (before end auction) End--")
	# 	elif choice == 3:
	# 		# End auction
	# 		wwalog.log("--End auction Start--")
	# 		wwautils.endAuction()
	# 		wwalog.log("--End auction End--")
	# 	else:
	# 		# Back
	# 		break
	#
	# return

def cmsMenu():
	while 1:
		choice = showMenu(["Back", "入交收list"])


		if choice == 1:
			# Check order list
			wwalog.log("--Check order list Start--")
			wwautils.checkOrderList()
			wwalog.log("--Check order list End--")



		else:
			# Back
			break

	return

def whatsappFunction():
	while 1:
		choice = showMenu(["Back", "確認交收訊息", "售後慰問"])
		if choice == 1:
			wwautils.meetingNotification()
		elif choice == 2:
			wwautils.PostOrder()
		else:
			break

# window = tk.Tk()
# window.title('Auction System')
# window.geometry('500x500')
# lbl_1 = tk.Label(window, text='Auction System')
# lbl_1.grid(column=0, row=0)
#
# #Auction Menu Button
# btn_1 = tk.Button(window, text='拍賣系統')
# btn_1['width']=50
# btn_1['height']=4
# btn_1.grid(column=0,row=1)
# btn_1['command'] = auctionMenu
#
# #CMS Menu
# btn_2 = tk.Button(window, text='CMS 網店操作')
# btn_2['width'] = 50
# btn_2['height'] = 4
# btn_2.grid(column=0, row=3)
# btn_2['command'] = cmsMenu
#
# #Whatsapp Menu
# btn_3 = tk.Button(window, text='Whatsapp 功能')
# btn_3['width'] = 50
# btn_3['height'] = 4
# btn_3.grid(column=0, row=2)
# btn_3['command'] = whatsappFunction
#
# window.mainloop()
