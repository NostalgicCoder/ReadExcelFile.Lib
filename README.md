-----------------------------------------
14th May 2024:
-----------------------------------------

The purpose of this project was to create a quick tool that I could use to extract data from a WPS Office spreadsheet file and into a format that could be imported easily into SQL Server Express. This code base will extract the data from multiple sheets of a spreadsheet file and write it out to a ‘.SQL’ file that can be run against the database.

Developer Limitations:

As I am using the free version of Visual Studio (Community Etd 2022) that comes with SQL Server Express I do not have access to the ETL import tools that comes with the paid for version of SSMS. I did a Google search and there did not appear to be a easy way to import data into SQL Express without investing in a paid for solution so this code base was used to overcome the problem.

Update 19/04/2024:

It appears if you install the SQL Server Management Studio add-on on top of a normal SQL Server Express installation this then gives you access to the ‘SQL Server Import & Export Wizard’ which I was not aware of at the time.

Technologies / Principles Used:

- C# (Console Test Harness & Class Library Project)
- EPPlus
- Microsoft.Office.Interop.Excel

Resources / Credits:

- https://coderwall.com/p/app3ya/read-excel-file-in-c
- https://stackoverflow.com/questions/44916744/do-i-need-to-have-office-installed-to-use-microsoft-office-interop-excel-dll

Images:

![Alt text](Images/ReadExcel1.JPG)
![Alt text](Images/ReadExcel2.JPG)
