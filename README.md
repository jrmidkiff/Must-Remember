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
**1. Opening and Renaming Client Data, Creating Main Categories and Moving 'Original Client Data' tables there**

To-Dos:
* Test the error-handler in Move_Tables if a table, say Applicants, is missing
* Test with older versions of Microsoft Excel
* Create a help file 

``` VBA
Option Compare Database
Global File_Location As String
Global original_client_data_catid As String

Function Client_Excel_Data()
Start:
    'File Location
        
    Dim Message, Title '(For testing, use P:\Test Data\VBA Test Data)
    Message = "Please enter the full file path for the client's excel data " & _
    "(no quotation marks, ending slashes '\', or file extensions)." & _
        Chr(13) & Chr(13) & _
        "The tables BOY, EOY, Hires, Promos, Terms, and Applicants must not already be defined in your database. If they are," & _
        " the import process will append the data to the pre-existing table." & _
        Chr(13) & Chr(13) & _
        "Keeping the excel file open will make copying field names into the module easier."
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
    
    Call Create_Categories 'This will run the Create_Categories subroutine, but it needs to start with "call"
    
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
       'The above is the command to actually import the data
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
    strSQL = "Alter Table " & table & " Add Column Auto_ID AutoIncrement"
    Debug.Print strSQL
    CurrentDb.Execute strSQL
     
    strSQL = "Alter Table " & table & " Add Constraint Auto_ID Primary Key(Auto_ID)"
    Debug.Print strSQL
    CurrentDb.Execute strSQL
   
End Sub

Function Create_Categories()
    str_insert = "INSERT INTO MSysNavPaneGroups (Flags, GroupCategoryID, Name, [Object Type Group], ObjectID, Position)"
    str_sql = str_insert & " VALUES (0, 3, 'Original Client Data', -1, 0, 0);"
    Debug.Print "The Category string is: "; str_final
    CurrentDb.Execute str_sql 'The table "MSysNavPaneGroups' controls all of the categories that appear in the database window.
    ' Adding in records to that table in the specified format will create these categories.
        
    original_client_data_catid = DLookup("[Id]", "MSysNavPaneGroups", "[Name] = 'Original Client Data'")
    Debug.Print "ID for the 'Original Client Data categories is: "; original_client_data_catid
        
    Move_Table ("BOY")
    Move_Table ("EOY")
    Move_Table ("Hires")
    Move_Table ("Promos")
    Move_Table ("Terms")
    Move_Table ("Applicants")
        
    str_sql = str_insert & " VALUES (0, 3, 'Final Client Data', -1, 0, 1);"
    Debug.Print str_final
    CurrentDb.Execute str_sql

    str_sql = str_insert & " VALUES (0, 3, 'Exclude', -1, 0, 2);"
    Debug.Print str_final
    CurrentDb.Execute str_sql

    str_sql = str_insert & " VALUES (0, 3, 'Appear/Disappear', -1, 0, 3);"
    Debug.Print str_final
    CurrentDb.Execute str_sql
    
    str_sql = str_insert & " VALUES (0, 3, 'Duplicates', -1, 0, 4);"
    Debug.Print str_final
    CurrentDb.Execute str_sql

    str_sql = str_insert & " VALUES (0, 3, 'Missing Race/Gender', -1, 0, 5);"
    Debug.Print str_final
    CurrentDb.Execute str_sql
    
    str_sql = str_insert & " VALUES (0, 3, 'Miscellaneous', -1, 0, 6);"
    Debug.Print str_final
    CurrentDb.Execute str_sql

    str_sql = str_insert & " VALUES (0, 3, 'Hires Applicants Matching', -1, 0, 7);"
    Debug.Print str_final
    CurrentDb.Execute str_sql

    Application.RefreshDatabaseWindow 'This refresh is necessary for the categories to appear. I think.
    DoCmd.NavigateTo ("Custom")
    
    
End Function

Function Move_Table(table)
    Dim table_object_id As String
    Debug.Print table
    On Error GoTo Error_Handler
    table_object_id = DLookup("[Id]", "MSysObjects", "[Name] = '" & table & "'")
    Debug.Print "Table is: "; table; " and Object_ID is: "; table_object_id
    
    str_move = "INSERT INTO MSysNavPaneGroupToObjects (Flags, GroupID, Icon, Name, ObjectID)"
    str_sql = str_move & " VALUES (0, " & original_client_data_catid & ", 0, '" & table & "', " & table_object_id & ");"
    Debug.Print str_sql
    
    
    CurrentDb.Execute str_sql
    
Error_Handler:
    If Err.Number <> 0 Then
        Debug.Print "Error Code: "; Err.Number; " Error Description: "; Err.Description
        Err.Clear
        Exit Function
    Else
        Exit Function
    End If
    
End Function
```

**2. Make Final and Exclude Tables**
To-Dos:

```VBA
Option Compare Database

Function Final_and_Exclude_Queries()
    prompt_response = MsgBox("This module will: " & Chr(13) & Chr(13) & _
        "1) Add Notes and Exclude fields to each table" & Chr(13) & _
        "2) Create and move the queries to make the Final and Exclude tables" & Chr(13) & Chr(13) & _
        "Would you like to run this module?", 68, "Final and Exclude Queries")
        '68 is the sum of 4 and 64 which correspond to vbYesNo and vbInformation respectively.
        'Search in AccessHelp for "MsgBox" for more info on this weird way of doing things
    If prompt_response = 6 Then       '6 = Yes, 7 = No
        'Add_Notes_and_Exclude ("BOY")
        'Add_Notes_and_Exclude ("EOY")
        'Add_Notes_and_Exclude ("Hires")
        'Add_Notes_and_Exclude ("Promos")
        'Add_Notes_and_Exclude ("Terms")
        'Add_Notes_and_Exclude ("Applicants")
        
        The_Queries ("BOY")
        The_Queries ("EOY")
        The_Queries ("Hires")
        The_Queries ("Promos")
        The_Queries ("Terms")
    
    MsgBox "The Final and Exclude Module has concluded. Check to be sure the queries have been set-up correctly (they have not been run). " & _
    "Each table will now have a field called 'NTL_Exclude' and 'NTL_Notes' at the end, which you should fill in throughout the validation process.", vbInformation, "Result"
    
    Else:
        Exit Function
    End If
    
    
End Function

Sub Add_Notes_and_Exclude(table)
    On Error Resume Next ' If an error occurs, this will silently ignore it and go to the
                         ' next statement
    strSQL_exclude = "Alter Table " & table & " Add Column NTL_Exclude Text;"
    strSQL_notes = "Alter Table " & table & " Add Column NTL_Notes Text;"
    
    'CurrentDb.Execute strSQL_exclude
    'CurrentDb.Execute strSQL_notes
End Sub

Function The_Queries(table)
    Dim qdf As DAO.QueryDef
    
    '1. Creating the final_table queries
    Set qdf = CurrentDb.CreateQueryDef("qryFinal_" & table)
    Application.RefreshDatabaseWindow
    
    str_sql_final = "SELECT " & table & ".* INTO tblFinal_" & table & _
              " FROM " & table & _
              " WHERE (((" & table & ".NTL_Exclude) Is Null));"
    Debug.Print str_sql_final
    qdf.SQL = str_sql_final
    
    '1A. Moving the final_table queries into "Final Client Data" category
    Dim Final_Client_Data_catid As String
    Final_Client_Data_catid = DLookup("[Id]", "MSysNavPaneGroups", "[Name] = 'Final Client Data'")
    Debug.Print "Destination_Category ID: "; Final_Client_Data_catid
    
    Call Move_This_Object("qryFinal_" & table, Final_Client_Data_catid)
    
    '2. Creating the exclude queries
    Set qdf = CurrentDb.CreateQueryDef("qryExclude_" & table)
    Application.RefreshDatabaseWindow
    
    str_sql_Exclude = "SELECT " & table & ".* INTO tblExclude_" & table & _
              " FROM " & table & _
              " WHERE (((" & table & ".NTL_Exclude) Is Not Null));"
    Debug.Print str_sql_Exclude
    qdf.SQL = str_sql_Exclude
    
    '2A. Moving the exclude queries into the "Exclude" category
    Dim Exclude_catid As String
    Exclude_catid = DLookup("[Id]", "MSysNavPaneGroups", "[Name] = 'Exclude'")
    Debug.Print "Destination_Category ID: "; Exclude_catid
    
    Call Move_This_Object("qryExclude_" & table, Exclude_catid)
    
End Function

Public Sub Move_This_Object(object, destination_category_id) 'This is the function that is being called from other modules to move queries and tables
    Dim table_object_id As String
        
    On Error GoTo Error_Handler
    object_object_id = DLookup("[Id]", "MSysObjects", "[Name] = '" & object & "'")
    Debug.Print "Object is: "; object; " and Object_ID is: "; object_object_id
    
    str_move = "INSERT INTO MSysNavPaneGroupToObjects (Flags, GroupID, Icon, Name, ObjectID)"
    str_sql = str_move & " VALUES (0, " & destination_category_id & ", 0, '" & object & "', " & object_object_id & ");"
    Debug.Print "Insert (aka movement) string: "; str_sql

    CurrentDb.Execute str_sql
    
Error_Handler:
    If Err.Number <> 0 Then
        Debug.Print "Error Code: "; Err.Number; " Error Description: "; Err.Description
        Err.Clear
        Exit Sub
    Else
        Exit Sub
    End If
End Sub
```

**2.5 Duplicates**
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
* Problem: How to account for multiple Hires/Promos/Terms
* Check for missing gender or race and conflicting gender/race

**5. Hires and Applicants Matching**
To-Dos:
* Jesus fuck there's a long road for me to get here
