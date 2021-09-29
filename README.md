#RAPackageTracker

## Built With
Google Sheets API
Google Drive API
Tinkter (Tk GUI Toolkit)

![image](https://user-images.githubusercontent.com/42818731/135283153-6e6136e2-ce48-4471-9875-8129d0471175.png)

Link to software Demo: https://www.youtube.com/watch?v=OOF_whNfEbs

## How to Use this Application
Type in the student name and optionally the delivery company (ex. Amazon) and signature/intials. For pickup input the student name, resident ID number and signature/intials. 

## Why was this product made? 

Working as an RA in New Jersey Institute of Technology's biggest hall, we would recieve hundreads of packages for students. These needed to be manually logged before we could return them to students. I quickly realized most of this process could be automated and went about doing it. 

## Code Highlights

Google's Sheet API has a massive limitation in that it can only perform 100 read or writes every 100 seconds. This meant that traditinal forms of cell checking would be too costly and cause the program to go over my allowed read/writes and get locked out. This was a big issue as a large part of updating the sheet was to find the last empty row. A unique solution I found to this issue was to write a VLOOKUP into a cell to find the last empty row, and then read from that cell. This turned a 200-1000 cell checking operation into a 2 cell checking operation. 

    packageSheet.update_cell(1,13,'=MATCH("@",ARRAYFORMULA(A4:A&"@"),0)+2')
    i = int(packageSheet.cell(1,13).value)

