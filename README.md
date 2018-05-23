# Must-Remember
This is code for things that I have to remember how to do for work that I am otherwise quite likely to forget


## SQL  
* Using CASE (equivalent to switch or multiple if-elses)  
``` SQL
SELECT COUNT(*), 
CASE 
    WHEN number_grade > 90 THEN 'A'
    WHEN number_grade > 80 THEN 'B'
    WHEN number_grade > 70 THEN 'C'
    ELSE 'F'
    END AS 'letter_grade'
FROM student_grades
GROUP BY letter_grade;
```
* Get the name of objects (or reports for ex.) out of a document
```SQL
SELECT MSysObjects.Name INTO tblTablesMoveAAP
FROM MSysObjects
WHERE Left([MSysObjects].[Name], 3) = "rpt";
```
## VBA
1. Opening and Renaming Client Data

To-Dos:
* Test with older versions of Microsoft Excel
* File not found error handler
* Add a primary key to each table
* Add Applicants
* Create options to skip particular table imports
* Create a help file 

```VBA
Sub Client_Excel_Data()

Dim Message, Title, File_Location
Message = "Please enter the full path for the client's excel data"
Title = "Open Client Data"

File_Location = InputBox(Message, Title)
Debug.Print File_Location

End Sub
```
