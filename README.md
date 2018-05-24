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
``` VBA
Option Compare Database
Global File_Location As String

Function Client_Excel_Data()
    'File Location
        
    Dim Message, Title '(For testing, use P:\Test Data\VBA Test Data)
    Message = "Please enter the full file path for the client's excel data (no quotes or ending slashes '\')"
    Title = "Open Client Data"
    File_Location = InputBox(Message, Title)
    
    Debug.Print File_Location
        
    BOY_Import

    MsgBox "Import Module has concluded. Check to be sure the records totals line up with the excel sheets. " & _
    "If there were any tables in the database with these table names already, the data was appended to them. You must therefore delete the table " & _
    "and re-run the macro." & Chr(13) & Chr(13) & _
    "If the client has many blank rows in their Excel data, these will appear in the tables.", vbInformation, "Result"

End Function
    
    'BOY
    
Sub BOY_Import()
        Debug.Print "File Location is: "; File_Location
        Message = "Please enter the exact BOY spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
               & "File Name: '" & File_Location & "'"
        Title = "BOY Import"
        
        BOY_Sheet_Name = InputBox(Message, Title)
        
        If BOY_Sheet_Name <> "Skip" Then
           DoCmd.TransferSpreadsheet acImport, 10, "BOY", File_Location, True, BOY_Sheet_Name & "!"
           'Create_Autonumber ("BOY") 'Not ready yet!
           On Error GoTo BOY_Error_Handler
           
        End If
        
        Debug.Print BOY_Sheet_Name
                
BOY_Error_Handler:
        MsgBox "The path for the file name and/or sheet name is invalid. Try again or attempt manual import.", vbCritical, "File Not Found"
        BOY_Import
                  
End Sub



        'EOY
        
Sub EOY_Import()

End Sub
         Message = "Please enter the exact EOY spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
                 & "File Name: '" & File_Location & "'"
         Title = "EOY Import"
         
         EOY_Sheet_Name = InputBox(Message, Title)
         
         If EOY_Sheet_Name <> "Skip" Then
             DoCmd.TransferSpreadsheet acImport, 10, "EOY", File_Location, True, EOY_Sheet_Name & "!"
             'Create_Autonumber ("EOY") 'Not ready yet!
             On Error GoTo EOY_Import
             
         End If
        
         Debug.Print EOY_Sheet_Name
        
        
        'Hires
        
Hires_Import:
         Message = "Please enter the exact Hires spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
                 & "File Name: '" & File_Location & "'"
         Title = "Hires Import"
         
         Hires_Sheet_Name = InputBox(Message, Title)
         
         If Hires_Sheet_Name <> "Skip" Then
             DoCmd.TransferSpreadsheet acImport, 10, "Hires", File_Location, True, Hires_Sheet_Name & "!"
             'Create_Autonumber ("Hires") 'Not ready yet!
             On Error GoTo Hires_Import
             
         End If
        
         Debug.Print Hires_Sheet_Name
        
        
        'Promos
        
Promos_Import:
         Message = "Please enter the exact Promos spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
                 & "File Name: '" & File_Location & "'"
         Title = "Promos Import"
         
         Promos_Sheet_Name = InputBox(Message, Title)
         
         If Promos_Sheet_Name <> "Skip" Then
             DoCmd.TransferSpreadsheet acImport, 10, "Promos", File_Location, True, Promos_Sheet_Name & "!"
             'Create_Autonumber ("Promos") 'Not ready yet!
             On Error GoTo Promos_Import
             
         End If
        
         Debug.Print Promos_Sheet_Name
        
        
        'Terms
        
Terms_Import:
         Message = "Please enter the exact Terms spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
                 & "File Name: '" & File_Location & "'"
         Title = "Terms Import"
         
         Terms_Sheet_Name = InputBox(Message, Title)
         
         If Terms_Sheet_Name <> "Skip" Then
             DoCmd.TransferSpreadsheet acImport, 10, "Terms", File_Location, True, Terms_Sheet_Name & "!"
             'Create_Autonumber ("Terms") 'Not ready yet!
             On Error GoTo Terms_Import
             
         End If
        
         Debug.Print Terms_Sheet_Name
        
        
        'Applicants
        
Applicants_Import:
         Message = "Please enter the exact Applicants spreadsheet name. Enter 'skip' to skip importing this file." & Chr(13) & Chr(13) _
                 & "File Name: '" & File_Location & "'"
         Title = "Applicants Import"
         
         Applicants_Sheet_Name = InputBox(Message, Title)
         
         If Applicants_Sheet_Name <> "Skip" Then
             DoCmd.TransferSpreadsheet acImport, 10, "Applicants", File_Location, True, Applicants_Sheet_Name & "!"
             'Create_Autonumber ("Applicants") 'Not ready yet!
             On Error GoTo Applicants_Import
             
         End If
        
         Debug.Print Applicants_Sheet_Name
    

   

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

```
