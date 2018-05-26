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
* Test with older versions of Microsoft Excel
* Add a primary key to each table
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
    "If there were any tables in the database with these table names already, the data was appended to them. You must therefore delete the table " & _
    "and re-run the macro." & Chr(13) & Chr(13) & _
    "If the client has many blank rows in their Excel data, these will appear in the tables.", vbInformation, "Result"

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
       'Create_Autonumber ("BOY") 'Not ready yet!
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
    'Create an AutoNumber called “Auto_ID” in specified table
    Dim db As DAO.Database
    Dim fld As DAO.Field
    Dim tdf As DAO.TableDef
    
    Set db = Application.CurrentDb
    Set tdf = db.TableDefs(table)
    ' First create a field with datatype = Long Integer
    Set fld = tdf.CreateField("Auto_ID", dbLong)
    With fld
    .Attributes = .Attributes Or dbAutoIncrField
    End With
    With tdf.Fields
    .Append fld
    .Refresh
    End With
End Sub



Sub Test()
Start:
    Dim Message, Title '(For testing, use P:\Test Data\VBA Test Data)
        Message = "Please enter the full file path for the client's excel data (no quotes or ending slashes '\')"
        Title = "Open Client Data"
        File_Location = InputBox(Message, Title)
        Debug.Print "Test File Location is: "; File_Location
        

End Sub


```
