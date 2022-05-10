from ast import parse
from http.client import SWITCHING_PROTOCOLS
import os
import configparser
from utils.parse import *
import time


rows=[]

OperatingSystem="windows"
ClearScreen="cls"
def runparser():

        rows=parsefile()        
        createExcel(rows)
        print ("")
        time.sleep(3)


        
        
        
def menu():	
	#options
        global ClearScreen
        os.system(ClearScreen)
        print ("Please choose an option below")
        print ("\t1 - Configure result location")
        print ("\t2 - Configure email consumers")
        print ("\t3 - View Log")
        print ("\t4 - Configure file to parse")
        print ("\t5 - Run")
        print ("\t6 - Exit")
 
def submenu1():
	#options
        global ClearScreen
        os.system(ClearScreen)
        print ("Please choose an option below")
        print ("\t1 - View result location")
        print ("\t2 - Replace result location")
        print ("\t3 - back")

def submenu2():
	#options
        global ClearScreen
        os.system(ClearScreen)
        print ("Please choose an option below")
        print ("\t1 - View stored emails")
        print ("\t2 - Replace emails")
        print ("\t3 - back")

def submenu4():
	#options
        global ClearScreen
        os.system(ClearScreen)
        print ("Please choose an option below")
        print ("\t1 - View path and name of raw file to parse")
        print ("\t2 - Replace path and name ")
        print ("\t3 - back")

def option1():
        while True:
                # show menu
                submenu1()
                                
                # request an option to user
                optionMenu = input("What wold you like to do? >> ")
        
                if optionMenu=="1":
                        read_config = configparser.ConfigParser()
                        read_config.read("app.config")
                        outputFile = read_config.get("DEFAULT", "locationexport")
                        print(outputFile+"Output.xlsx")
                        time.sleep(2)
                elif optionMenu=="2":
                        print ("")
                        location=input("Please enter a valid path to export output file and ended with /  ")
                        
                        saveproperty("locationexport",location)
                
                elif optionMenu=="3":
                        break
                else:
                        print ("")
                        input("Not valid choice try again")

def option2():
        while True:
                # show menu
                submenu2()
                                
                # request an option to user
                optionMenu = input("What wold you like to do? >> ")
        
                if optionMenu=="1":
                        read_config = configparser.ConfigParser()
                        read_config.read("app.config")
                        Email = read_config.get("DEFAULT", "Email")
                        print(Email)
                        time.sleep(2)
                elif optionMenu=="2":
                        print ("")
                        location=input("Please enter a valid email(s) delimited by comma  ")
                        
                        saveproperty("Email",location)
                
                elif optionMenu=="3":
                        break
                else:
                        print ("")
                        input("Not valid choice try again")                      

def option3():

        file1 = open('std.log', 'r')
        Lines = file1.readlines()
        
        count = 0
        # Strips the newline character
        for line in Lines:
                count += 1
                print("Line{}: {}".format(count, line.strip()))
        input("key press to continue")  

def option4():
        while True:
                # show menu
                submenu4()
                                
                # request an option to user
                optionMenu = input("What wold you like to do? >> ")
        
                if optionMenu=="1":
                        read_config = configparser.ConfigParser()
                        read_config.read("app.config")
                        outputFile = read_config.get("DEFAULT", "locationfile")
                        print(outputFile)
                        time.sleep(2)
                elif optionMenu=="2":
                        print ("")
                        location=input("Please enter a valid path and name of the raw file  ")                        
                        saveproperty("locationfile",location)
                
                elif optionMenu=="3":
                        break
                else:
                        print ("")
                        input("Not valid choice try again")

def saveproperty(propname,value):
        try:
        
                read_config = configparser.ConfigParser()
                read_config.read("app.config")
                locationexport = read_config.get("DEFAULT", "locationexport")
                locationfile = read_config.get("DEFAULT", "locationfile")
                OperatingSystem = read_config.get("DEFAULT", "OperatingSystem")
                Email = read_config.get("DEFAULT", "Email")

                if propname == "locationexport":
                        locationexport=value
                elif propname == "locationfile":
                        locationfile=value
                elif propname == "Email":
                        Email=value
                else:
                        print("not a property")

                config = configparser.ConfigParser()
                config['DEFAULT'] = {'locationexport': locationexport,
                                'locationfile':locationfile,
                                'OperatingSystem': OperatingSystem,
                                'Email':Email}

                with open('app.config', 'w') as configfile:
                        config.write(configfile)
                        
                print("Property saved")
                time.sleep(2)


        except:
                print("Property not saved")
                time.sleep(3)
                               
while True:
	# show menu
	menu()
                         
	# request an option to user
	optionMenu = input("What wold you like to do? >> ")
 
	if optionMenu=="1":
		option1()
	elif optionMenu=="2":
		option2()
	elif optionMenu=="3":
		option3()
	elif optionMenu=="4":
                option4()
	elif optionMenu=="5":
                runparser()              
	elif optionMenu=="6":
		break
	else:
		print ("")
		input("Not valid choice try again")


