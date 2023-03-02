# Password saver

A small script for saving user personas for the test environment of an application
to the clipboard. The application will read from an excel file and takes username and
password combinations. Only accepts ".xlsx" files for now

run for more help:

`python3 <path-to-pwdsaver.py> -h`

## Install

1. Download the latest release
2. `pip install -r requirements.txt`
3. Add a 'personas' sheet to your Excel file with the username
    / password combinations. 

## To run
`python3 <path-to-pwdsaver.py> [FILE]` 