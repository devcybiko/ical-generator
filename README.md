# iCal Generator
- originally from https://github.com/flowor/ical-generator.git
- converts .csv to .ics files

## Files
- calendar.csv - a csv file for conversion
- calendar.ics - the resultant .ics file (from .csv)
- csv-to-ics.bat - a windows.bat file to run the app (for automation)
- csv-to-ics.js - the main program (extended from script.js)
- csv-to-ics.vbs - visual basic script to run csv-to-ics.js (for automation)
- index.html - the original HTML driver
- macro.vba - the Outlook VBA macro that reads the calendar and outputs a .csv file
- outlook.png - image for the original HTML driver
- README-orig.md -  from original FLOWOR repo
- README.md - this file
- sample_calendar.csv - original example.csv
- script.js - original JS 
- style.css - CSS for original HTML driver

## Running
- node csv--to-ics.js <infile.csv> <outfile.ics>
- infile.csv = the input CSV (default: calendar.csv)
- outfile.ics = the output .ics (iCal) file

## VBA Macro

The VBA Macro is added to Outlook by...

- Running Outlook
- Selecting the Calendar
- ALT-F11 (brings up the VBA Macro Editor)
- copy in "macro.vba"
- Run it with F5

The VBA Macro will:

- read 14 days of events
- package it into a calendar.csv
- run csv-to-ics.js
- send an email with attachment (calendar.ics) to me

Sadly, the macro cannot easily be scheduled to run periodically. So I must remember to run it by hand daily.

## CSV Format

The file should contain the following headers in any order: 

* Subject
* Start Date
* Start Time
* Date Stamp
* End Date
* End Time
* All Day
* Description
* Location
* UID
* Busy Status

### Busy Status

_Busy Status_ can use the following keywords:

* FREE
* WORKINGELSEWHERE
* TENTATIVE
* BUSY
* AWAY# ical-generator
