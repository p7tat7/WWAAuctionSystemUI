# -*- coding: UTF-8 -*-

import wwaglobal
import os
from datetime import datetime

f = None

def log(msg):
	global f
	f = open("logs\\log-" + wwaglobal.today + ".txt", "a", encoding = "utf-8")
	f.write(str(msg) + "\n")
	f.close()

def createLog():
	global f

	currentTime = datetime.today().strftime('%H:%M:%S')
	logFile = "logs\\log-" + wwaglobal.today + ".txt"

	if not os.path.isdir("logs"):
		# create logs folder
		os.mkdir("logs")

	if os.path.isfile(logFile):
		# log file already exists
		f = open(logFile, "a", encoding = "utf-8")
		f.write("\n\n\n\n\n-----" + wwaglobal.today + " " + currentTime + "-----\n\n")
		f.close()
	else:
		# log file does not exist
		f = open(logFile, "w+", encoding = "utf-8")
		f.write("-----" + wwaglobal.today + " " + currentTime + "-----\n\n")
		f.close()

def closeLog():
	global f
	f.close()
