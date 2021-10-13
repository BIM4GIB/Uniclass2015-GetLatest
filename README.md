#########################################################################################
#   ________  _______   ___      ___ ___  _________  _______   _______   ________       #  
#  |\   __  \|\  ___ \ |\  \    /  /|\  \|\___   ___\\  ___ \ |\  ___ \ |\   __  \      #  
#  \ \  \|\  \ \   __/|\ \  \  /  / | \  \|___ \  \_\ \   __/|\ \   __/|\ \  \|\  \     #  
#   \ \   _  _\ \  \_|/_\ \  \/  / / \ \  \   \ \  \ \ \  \_|/_\ \  \_|/_\ \   _  _\    #  
#    \ \  \\  \\ \  \_|\ \ \    / /   \ \  \   \ \  \ \ \  \_|\ \ \  \_|\ \ \  \\  \|   #  
#     \ \__\\ _\\ \_______\ \__/ /     \ \__\   \ \__\ \ \_______\ \_______\ \__\\ _\   #  
#      \|__|\|__|\|_______|\|__|/       \|__|    \|__|  \|_______|\|_______|\|__|\|__|  # 
#########################################################################################
# 
#######################
# Uniclass2015 Tables #
#ClassificationManager#
# GetLatest&Merge     #
# v0.10               #
# 09/08/19            #
# by RPG @BIM4GIB     #
# reviteer@hotmail.com#
#######################
#
############################################################################################################################################################
#rev v0.2 bug fixes (i.e. it actually runs now)
#rev v0.3 updated order of tables, following table PM v1.0 release
#rev v0.4 added Classification Manager Custom Database UK-Uniclass2015.xlsx with data connections to Uniclass2015-AllTables.xlsx
#rev v0.5 added dialog box to confirm script run successfully and added autoupdating of the Classification Manager Database
#rev v0.6 fixed form (dialog box). Now it does display even when not running in IDE.
#rev v0.7 added Roles table, added flexibility to run regardless of location in the local computer
#rev v0.8 Temporarily disabled the classification manager database, as it needs some attention 
#rev v0.9//2019.06.19// So much better now. Excel doesn't open while script is working, got the Classification Manager Database updater back in biz...So gud
#rev v0.10//2019.08.09 NBS changed their website, broke script but now fixed+works a bit faster. Results window comes into focus now. 
#           Added a line to force use of TLS1.2 to avoid problems with TLS1.1. Now downloads PDFs and place in folder named YYMM
#rev v0.11//2021.10.13// Now in GitHub! Also added -UseBasicParsing parameter for when IE is not present/initialised
#
#
#TODO: *Put some loops in there, make it pretty, add some classes and functions too
#      
#############################################################################################################################################################

############################################################################################################################################
#WHAT?#	  
#######
The purpose of this tool is to download the latest 2015 tables from the NBS website and merge them into a single spreadsheet.
It also includes a Uniclass 2015 custom database for the Classification Manager add-in for Autodesk Revit, which will update itself
automatically after the script is run (ClassificationManagerDatabase-Uniclass2015.xlsx).
############################################################################################################################################

############################################################################################################################################
#HOW? #
#######
 1. If you downloaded from the interwebs, copy the folder "Uniclass2015-GetLatest" from the ZIP file to e.g. your desktop, and open it.
 2. Double-click on the "Uniclass2015-GetLatest" shortcut. A PowerShell icon will appear in your taskbar. Click 'Open' on the dialog box.
    You might have to type "Y" (without the quotes) in PowerShell to confirm you want to run the script, depending on your security settings.
 3. Wait approx. 20s.
 4. After the script has run, a dialog box will appear with some useless stats. Click OK.
 5. You should now see a new spreadsheet called "Uniclass2015-AllTables.xlsx" inside the same folder as the shortcut that called the script.
    This can be used for reference, as a look-up for other tools, etc.
 6. The script will also update 'ClassificationManagerDatabase-Uniclass2015.xlsx' which can then be used as a custom database with the
    Classification Manager addin for Revit (https://www.biminteroperabilitytools.com/classificationmanager.php)
############################################################################################################################################

############################################################################################################################################
#WHY! #
#######
If you get errors or nothing happens:
  a. Navigate to the "Resources" folder
  b. Right-click on "Uniclass2015-GetLatest.ps1"
  c. Select "Properties"
  d. On the General tab, make sure you Unblock the file
  e. Click OK, then try again
  f. If you still have problems, please get in touch!

############################################################################################################################################

############################################################################################################################################
#WHO? #
#######
COPYLEFT:

   -Uniclass 2015 is a unified classification system by NBS (www.theNBS.com) 
    and is licensed for use under the terms of the Creative Commons Attribution-NoDerivatives 4.0 International licence.

    This "program" is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation.
    https://creativecommons.org/licenses/by-nd/4.0/

    This "program" is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
    See the GNU General Public License for more details.

    See <https://www.gnu.org/licenses/gpl.txt>.
