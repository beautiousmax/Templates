# Templates

The goal of this program is to reformat the raw turning-movement data collected by Traffic Data Online (or by other traffic engineers) into a customized report to give to clients. The program was designed to process many data files at once, with minimal user interaction.


## Requirements

Python

Python Image Library

openpyxl

xml.etree.ElementTree



## How To Use

Download and build using pyinstaller.

Drag and drop Traffic Data Online excel files onto the exe, or double click on it. A window will appear. The minimum information required to build a report is a date, a start time and end time in the 'Time Frame' section, and at least one box checked. The time can either be in a 24 hour or a 12 hour format, for example "9:00 pm" or "21:00". If the program can't process the time you enter, it should tell you why not.

If you drag and drop Traffic Data Online excel files, most of the information will be entered automatically into the window, such as the date, times, and the boxes that should be checked. The program will also guess at labeling the legs of the intersection the report is for. If you are processing multiple files at a time, you can use buttons to scroll through each file to edit the header information. When the information is correct, click on 'Create'. Your finished reports will be added to the directory your exe is in, ready for review. If you did not use Traffic Data Online excel files, you will need to manually enter in the missing data.



Copyright 2015 Clara Dawn Griffith