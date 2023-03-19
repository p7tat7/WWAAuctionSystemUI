# -*- coding: UTF-8 -*-

from datetime import datetime
import os
import sys
import shutil


VERSION = "1.2"

today = ""
dropboxPath = ""
personalDropboxPath = ""
auctionDocumentPath = ""
photoPath = ""
chromeDriverPath = ""
chromeProfilePath = ""
chromeProfilePath1 = ""
chromeProfilePath2 = ""
src_placement = ''
src_path = ''
grp_content_csv_path = ''
debug = False







def init():
	global today
	global dropboxPath
	global personalDropboxPath
	global auctionDocumentPath
	global photoPath
	global chromeDriverPath
	global chromeProfilePath
	global chromeProfilePath1
	global chromeProfilePath2
	global debug
	global src_placement
	global src_photoPath
	global src_path
	global grp_content_csv_path

	today = datetime.today().strftime('%Y-%m-%d')

	if len(sys.argv) > 1:
		debug = sys.argv[1] == "debug"

	if os.path.isdir("C:\\Users\\world\\World Wide Aquarium Dropbox\\WWA SHKS\\"):
		# WWA desktop
		dropboxPath = "C:\\Users\\world\\World Wide Aquarium Dropbox\\WWA SHKS\\"
		personalDropboxPath = "C:\\Users\\world\\World Wide Aquarium Dropbox\\SHKS WWA\\"
		chromeDriverPath = "C:\\Users\\world\\OneDrive\\文件\\chromedriver.exe"
		chromeProfilePath = r"C:\Users\world\AppData\Local\Google\Chrome\User Data"
	elif os.path.isdir("C:\\Users\\tommy\\World Wide Aquarium Dropbox\\WWA SHKS\\"):
		# Tommy laptop
		dropboxPath = "C:\\Users\\tommy\\World Wide Aquarium Dropbox\\WWA SHKS\\"
		personalDropboxPath = "C:\\Users\\tommy\\World Wide Aquarium Dropbox\\SHKS WWA\\"
		chromeDriverPath = "C:\\Users\\tommy\\chromedriver.exe"
		chromeProfilePath = "C:\\Users\\tommy\\AppData\\Local\\Google\\Chrome\\User Data"
	elif os.path.isdir("C:\\Users\\user\\World Wide Aquarium Dropbox\\WWA SHKS\\"):
		# Ryan desktop
		dropboxPath = "C:\\Users\\user\\World Wide Aquarium Dropbox\\WWA SHKS\\"
		personalDropboxPath = "C:\\Users\\user\\World Wide Aquarium Dropbox\\SHKS WWA\\"
		chromeDriverPath = "C:\\Users\\user\\Documents\\chromedriver.exe"
		chromeProfilePath = "C:\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data"
	elif os.path.isdir("C:\\Users\\p7tat7\\World Wide Aquarium Dropbox\\WWA SHKS\\"):
		# Ryan laptop
		dropboxPath = "C:\\Users\\p7tat7\\World Wide Aquarium Dropbox\\WWA SHKS\\"
		personalDropboxPath = "C:\\Users\\p7tat7\\World Wide Aquarium Dropbox\\SHKS WWA\\"
		chromeDriverPath = "C:\\Users\\p7tat7\\Documents\\chromedriver.exe"
		chromeProfilePath = "C:\\Users\\p7tat7\\AppData\\Local\\Google\\Chrome\\User Data"
	else:
		# unknown computer
		dropboxPath = input("Path of dropbox \"WWA SHKS\" folder (...\\WWA SHKS\\): ")
		personalDropboxPath = input("Path of dropbox \"SHKS WWA\" folder (...\\SHKS WWA\\): ")
		chromeDriverPath = input("Chrome driver path (...\\chromedriver.exe): ")
		chromeProfilePath = input("Chrome profile path (...\\Google\\Chrome\\User Data): ")


	chromeProfilePath1 = chromeProfilePath + "1"
	chromeProfilePath2 = chromeProfilePath + "2"

	auctionDocumentPath = dropboxPath + datetime.today().strftime('%Y')+ "\\Auction Coral Retial\\" + today + " Auction Coral\\品種.docx"
	grp_content_csv_path = dropboxPath + datetime.today().strftime('%Y')+ "\\Auction Coral Retial\\" + today + " Auction Coral\\" + today + 'Photo Group.csv'

	photoPath = dropboxPath + datetime.today().strftime('%Y')+ "\\Auction Coral Retial\\" + datetime.today().strftime('%Y-%m-%d') + " Auction Coral\\Picture"
	src_path = dropboxPath + "PROGRAM\\AuctionSystem UI\\src\\"
	src_photoPath = dropboxPath + "PROGRAM\\AuctionSystem UI\\src\\Picture"
	src_placement = dropboxPath + "PROGRAM\\AuctionSystem UI\\src\\Placement Photo"
	# print("Copying Chrome profile.")
	# #wwalog.log("Copying Chrome profile.")
	# os.system(
	# 	"Xcopy \"" + chromeProfilePath + "\" \"" + chromeProfilePath1 + "\" /E/H/C/I/Y/Q > nul 2>&1")

def isInt(str):
	try:
		int(str)
		return True
	except:
		return False
