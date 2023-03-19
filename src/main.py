# -*- coding: UTF-8 -*-

import wwaglobal
import wwalog
import wwamenu
import wwautils
import os
import sys
import wwatkmenu
import shutil
import subprocess

os.system("title WWA Auction System v" + wwaglobal.VERSION)
os.system("cls")
print("WWA Auction System v" + wwaglobal.VERSION)
print()

# init global variables
wwaglobal.init()

# create log file
wwalog.createLog()

if wwaglobal.debug:
	print("[DEBUG MODE]")
print("Copying Chrome profile.")
if not wwaglobal.debug:
	# try:
	# 	shutil.copytree(wwaglobal.chromeProfilePath, wwaglobal.chromeProfilePath1)
	# except:
	# 	shutil.rmtree(wwaglobal.chromeProfilePath1)
	# 	shutil.copytree(wwaglobal.chromeProfilePath, wwaglobal.chromeProfilePath1)
	#os.system('copy ' + wwaglobal.chromeProfilePath +' ' + wwaglobal.chromeProfilePath1)
	source = wwaglobal.chromeProfilePath
	target = wwaglobal.chromeProfilePath1
	#print(wwaglobal.chromeProfilePath)
	#os.system("""xcopy "%s" "%s" C:\Windows\system32""" % (source, target))
	#print(os.system("Xcopy \"" + wwaglobal.chromeProfilePath + "\" \"" + wwaglobal.chromeProfilePath1 + "\" /E/H/C/I/Y/Q > nul 2>&1"))
	#print(os.system("Xcopy %s %s" % (src, dst)))
	#subprocess.call(['xcopy', src, dst])
	#os.system("Xcopy %s %s" % (src, dst))
	os.system("Xcopy \"" + wwaglobal.chromeProfilePath + "\" \"" + wwaglobal.chromeProfilePath1 + "\" /E/H/C/I/Y/Q > nul 2>&1")
	#shutil.copytree(wwaglobal.chromeProfilePath, wwaglobal.chromeProfilePath1, dirs_exist_ok=True)
# show menu
wwatkmenu.main()

wwalog.log("Program end.")
