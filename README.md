# Mass-Printer-Install-Script
Powershell script to install Ricoh or HP printers on multiple computers from a list. 

This script was created as a method to mass install printers on a list of networked devices. The goal of this project was to provide an easy and methodical system to install printers in mass. Some possible improvements down the line is to take all printers that should be installed at the start, but for my purposes I only wanted it to install one printer at a time. Results of the installs are saved and output to a file for review to see which devices the printers did not install on, and to get a report of which drivers were used.

## To use
Create a this file C:\Temp\PRINTER_computers.txt
This file will contain a list of all computers the printers should be added to - one computer name per line
Run the script and follow the prompts. 

This script allows for hard coded printer drivers, or the ability for the script to pick the most up to date driver on the device. The hard coded option is the fastest to use, but in some cases where the installed drivers are unknown, the "Most Updated" option may be more suitable.

At the end of the install for each printer, the script will ask if you want to repeat the installer for additional printers, or end the script. When you choose to end the script, the output log file will be saved at C:\Temp\printers_MMDDYY.csv.
