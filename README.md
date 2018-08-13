# create_outlook_folders
Script used to create outlook folders programatically

At work I have a large amount of mailing lists I am subscribed to. In order to keep these all managed I have them in folders. Managing this folder structure using the GUI has been time consuming and it is something all employees have to do. This script is written to create the following folder structure:

Inbox
-Mailing Lists
--Sub Folder 1
--Sub Folder 2
--Etc

The script will first check to see if the Inbox\Mailing Lists folder exists, if not create it. Then it will iterate through the mailing_lists_array, check if each folder exists and if not create it. This allows for future updates to the subfolder list without causing any issues or modifying folders already present.

Confirmed working on:
Outlook 2016

Download the ZIP file, extract it somewhere. Open powershell, change to the directory you just extracted the files to and run the script.

Please note this isn't signed, so you will have to allow unsigned scripts:
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-6
