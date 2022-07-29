# Resident Assistant (RA) Package Tracker


![image](https://user-images.githubusercontent.com/42818731/181652907-7e0d8893-18fd-43c0-b592-fbfece1ab8dc.png)



## Description

This Python application looks up names from a CSV student roster to collect room, hall, floor and email information. Then completes the sheet with date of logging as well as all that information and sends an alert email to the student's email address. It also assigns the order a package number (each floor has different package order numbering), which will be put on the package for tracking and storing. After the student picks up their package, it again finds the correct order and updates the date of pickup and staff member who delivered the package. 

This automates a process that used to take 2 minutes per package to do manually to less than 5 seconds per package.

![Screen Shot 2021-09-30 at 10 27 51 AM](https://user-images.githubusercontent.com/42818731/135475192-f88ef58b-fa2b-4b88-9681-a9eaf08e6672.jpg)

Link to software Demo: https://www.youtube.com/watch?v=OOF_whNfEbs

## Built With
* Google Sheets API
* Google Drive API
* Tinkter (Tk GUI Toolkit)

## How to Use this Application
Type in the student name and optionally the delivery company (ex. Amazon) and signature/intials. For pickup input the student name, resident ID number and signature/intials. 

## Why was this product made? 

Working as an RA in New Jersey Institute of Technology's biggest hall, we would recieve hundreads of packages for students. These needed to be manually logged before we could return them to students. I quickly realized most of this process could be automated and went about doing it. 

## Code Highlights

Google's Sheet API has a massive limitation in that it can only perform 100 read or writes every 100 seconds. This meant that traditional forms of cell checking would be too costly and cause the program to go over my allowed read/writes and get locked out. This was a big issue as a large part of updating the sheet was to find the last empty row. A unique solution I found to this issue was to write a VLOOKUP into a cell to find the last empty row, and then read from that cell. This turned a 200-1000 cell checking operation into a 2 cell checking operation. 

    packageSheet.update_cell(1,13,'=MATCH("@",ARRAYFORMULA(A4:A&"@"),0)+2')
    i = int(packageSheet.cell(1,13).value)

