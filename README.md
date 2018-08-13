# create_outlook_folders
Script used to create outlook folders programatically

At work I have a large amount of mailing lists I am subscribed to. In order to keep these all managed I have them in folders. Managing this folder structure using the GUI has been time consuming and it is something all employees have to do. This script is written to create the following folder structure:

Inbox
-Mailing Lists
--Sub Folder 1
--Sub Folder 2
--Etc

The script will iterate through the array, check if each folder exists and if not create it. This allows for future updates to the subfolder list without causing any issues or modifying folders already present.

Confirmed working on:
Outlook 2016
