# excel2csv
Excel to CSV conversion tool

This program is a simple and easy to use tool to convert Excel files in  xls and xlsx formats to csv files.

The tool will export each sheet to a csv file. You can also pass sheets names that needed to be converted only along with corresponding output csv files names.
Its deployed as fat Jar file.

It can convert Excel files effeciently and enhance the conversion by avoiding common conversion issues such as numbers representation as exponential number and also if data column contains ',' character.

This tool and its source code is availbe for free and commertial applications usage.

Usage examples:

Ex1: java -jar excel2csv.jar sample.xlsx /home
Ex2: java -jar excel2csv.jar sample.xls C:\home sheet1,sheet2,sheet3
Ex3: java -jar excel2csv.jar sample.xlsx /home sheet1,sheet2,sheet3 AA,BB,CC
