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
**1. Opening and Renaming Client Data**

To-Dos:
* Insert note in first dialogue box recommending to keep excel file open for easy name copy-pasting
* Test with older versions of Microsoft Excel
* Create a help file 
``` VBA
Option Compare Database
Global File_Location As String

Function Client_Excel_Data()
Start:
    'File Location
        
    Dim Message, Title '(For testing, use P:\Test Data\VBA Test Data)
    Message = "Please enter the full file path for the client's excel data " & _
    "(no quotation marks, ending slashes '\', or file extensions)." & _
        Chr(13) & Chr(13) & _
        "The tables BOY, EOY, Hires, Promos, Terms, and Applicants must not already be defined in your database. If they are," & _
        " the import process will append the data to the pre-existing table."
    Title = "Open Client Data"
    File_Location = InputBox(Message, Title)
    
    If File_Location = "" Then 'If someone wants to exit the main dialogue, bring up a box to let them
        exit_response = MsgBox("Would you like to end the import process?", vbYesNo, "Exit")
        If exit_response = 6 Then       '6 = Yes, 7 = No
            Exit Function
        Else:
            GoTo Start
        End If
    End If
    Debug.Print File_Location
        
    import_subroutine ("BOY") 'Run the import subroutine for BOY
    import_subroutine ("EOY") 'Run the import subroutine for EOY, etc. etc.
    import_subroutine ("Hires")
    import_subroutine ("Promos")
    import_subroutine ("Terms")
    import_subroutine ("Applicants")

    MsgBox "Import Module has concluded. Check to be sure the records totals line up with the excel sheets. " & _
    "If there were any tables in the database with these table names already, the data was appended to them and you probably do not want that." & _
    Chr(13) & Chr(13) & _
    "There is now a Primary Key Auto Number field at the end of each table called 'Auto_ID'.", vbInformation, "Result"

End Function
Sub import_subroutine(table)
Start:
    Message = "Please enter the exact " & table & " spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
           & "File Name: '" & File_Location & "'"
    Title = table & " Import"
    
    Sheet_Name = InputBox(Message, Title)
    Debug.Print "Table name is: "; table
    If Sheet_Name = "skip" Then
       Exit Sub
    ElseIf Sheet_Name = "" Then 'If someone wants to exit the main dialogue, they can
        exit_response = MsgBox("Would you like to end the import process?", vbYesNo, "Exit")
        Debug.Print "Exit Response is: "; exit_response
        If exit_response = 6 Then       '6 = Yes, 7 = No
            End
        Else:
            GoTo Start
        End If
    Else
       On Error GoTo Error_Handler:
       DoCmd.TransferSpreadsheet acImport, 10, table, File_Location, True, Sheet_Name & "!"
       Create_Autonumber (table) 'Use the Create_Autonumber sub to create the field 'Auto_ID
                                 'and have that field set as the primary key
    End If
    
    Debug.Print "Sheet name is: "; Sheet_Name
            
Error_Handler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number; Err.Description
        MsgBox "The path for the file name and/or sheet name is invalid. Try again or attempt manual import.", vbCritical, "File Not Found"
        Resume Start
    Else
        Exit Sub
    End If
End Sub
   
Sub Create_Autonumber(table)
    If table <> "Applicants" Then
        strSQL = "Alter Table " & table & " Add Column Auto_ID AutoIncrement, " & _
        "ERace Text(100), EGender Text(100)"
    Else
        strSQL = "Alter Table " & table & " Add Column Auto_ID AutoIncrement"
    End If
    Debug.Print strSQL
    CurrentDb.Execute strSQL
     
    strSQL = "Alter Table " & table & " Add Constraint Auto_ID Primary Key(Auto_ID)"
    Debug.Print strSQL
    CurrentDb.Execute strSQL
   
End Sub

```
**2. Make Final and Exclude Tables**
To-Dos:

**3. Appear/Disappear**
To-Dos:
* Be sure to keep original client field names intact
* Ask for EmpID and save it for each table

**4. Missing Gender/Race**
To-Dos:
* Bring ERace/EGender field creation from Module 1 to this one
* Ask for Race/Gender field names
* Create tblERace
* Update ERace/EGender in each table
* Note about recommending to re-run "Final" and "Exclude" mktable queries
* Make full table of all employees
* Check for missing gender or race and conflicting gender/race
