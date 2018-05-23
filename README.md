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
* File not found error handler
* Add a primary key to each table
* Add Applicants
* Create options to skip particular table imports
* Create a help file 

Option Compare Database

Sub Client_Excel_Data()
    'File Location
    Dim Message, Title, File_Location '(For testing, use P:\Test Data\VBA Test Data)
    Message = "Please enter the full file path for the client's excel data (no quotes or ending slashes '\')"
    Title = "Open Client Data"
    
    File_Location = InputBox(Message, Title)
    Debug.Print File_Location
    
    'BOY
    Message = "Please enter the exact BOY spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
            & "File Name: '" & File_Location & "'"
    Title = "BOY Import"
    
    BOY_Sheet_Name = InputBox(Message, Title)
    
    If BOY_Sheet_Name <> "Skip" Then
        DoCmd.TransferSpreadsheet acImport, 10, "BOY", File_Location, True, BOY_Sheet_Name & "!"
        'Create_Autonumber ("BOY") 'Not ready yet!
    End If
   
    Debug.Print BOY_Sheet_Name
    
      
    'EOY
    Message = "Please enter the exact EOY spreadsheet name" & Chr(13) & Chr(13) _
            & "File Name: '" & File_Location & "'"
    Title = "EOY Import"
    
    EOY_Sheet_Name = InputBox(Message, Title)
    Debug.Print EOY_Sheet_Name
    
    If BOY_Sheet_Name <> "Skip" Then
        DoCmd.TransferSpreadsheet acImport, 10, "EOY", File_Location, True, EOY_Sheet_Name & "!"
        'Create_Autonumber ("BOY") 'Not ready yet!
    End If
    
    
    'Hires
    Message = "Please enter the exact Hires spreadsheet name" & Chr(13) & Chr(13) _
            & "File Name: '" & File_Location & "'"
    Title = "Hires Import"
    
    Hires_Sheet_Name = InputBox(Message, Title)
    Debug.Print Hires_Sheet_Name
    
    If BOY_Sheet_Name <> "Skip" Then
        DoCmd.TransferSpreadsheet acImport, 10, "Hires", File_Location, True, Hires_Sheet_Name & "!"
        'Create_Autonumber ("BOY") 'Not ready yet!
    End If
        
        
    'Promos
    Message = "Please enter the exact Promos spreadsheet name" & Chr(13) & Chr(13) _
            & "File Name: '" & File_Location & "'"
    Title = "Promos Import"
    
    Promos_Sheet_Name = InputBox(Message, Title)
    Debug.Print Promos_Sheet_Name
    
    If BOY_Sheet_Name <> "Skip" Then
        DoCmd.TransferSpreadsheet acImport, 10, "Promos", File_Location, True, Promos_Sheet_Name & "!"
        'Create_Autonumber ("BOY") 'Not ready yet!
    End If
    
    
    'Terms
    Message = "Please enter the exact Terms spreadsheet name" & Chr(13) & Chr(13) _
            & "File Name: '" & File_Location & "'"
    Title = "Terms Import"
    
    Terms_Sheet_Name = InputBox(Message, Title)
    Debug.Print Terms_Sheet_Name
    
    If BOY_Sheet_Name <> "Skip" Then
        DoCmd.TransferSpreadsheet acImport, 10, "Terms", File_Location, True, Terms_Sheet_Name & "!"
        'Create_Autonumber ("BOY") 'Not ready yet!
    End If
    
    
    MsgBox "All files successfully imported. Check to be sure the records totals line up with the excel sheets. " & _
        "If there were any tables in the database with these table names already, the data was appended to them. You must therefore delete the table " & _
        "and re-run the macro." & Chr(13) & Chr(13) & _
        "If the client had left a lot of blank rows in Excel, these will appear in the tables. Remain calm.", vbInformation, "Result"
    
    
    
End Sub

Sub Create_Autonumber(Table)
    'Create an AutoNumber called “Auto_ID” in specified table
    Dim db As DAO.Database
    Dim fld As DAO.Field
    Dim tdf As DAO.TableDef
    
    Set db = Application.CurrentDb
    Set tdf = db.TableDefs(Table)
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
File_Location = "aentsh"
Message = "Hi" + Chr(13) + "Bye" + File_Location
Debug.Print Message

End Sub

