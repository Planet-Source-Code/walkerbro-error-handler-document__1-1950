<div align="center">

## Error Handler Document


</div>

### Description

This code pastes into a Module that Create (if not exists) a MDB to record the errors that occur in your application.
 
### More Info
 
Needs (DatabaseName, Date, Err.Number, Err.Description, PrivateNotes, Optional(User))

Load in "References" the "Microsoft DAO 3.51 Object Library"

Basic Error handling information.

True or False if it was succesful.

No known side effects.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[WalkerBro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/walkerbro.md)
**Level**          |Unknown
**User Rating**    |6.0 (615 globes from 103 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/walkerbro-error-handler-document__1-1950/archive/master.zip)

### API Declarations

```
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
  End Type
  Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
```


### Source Code

```
'*   Created by Walker Brother (tm)
'*   web page : http://www.walkerbro.8m.com
'*   e-mail  : info@walkerbro.8m.com
'*   This Module Logs the Errors your application may incounter into a MDB, if the MDB
'*   does not exist the it Creates it.
'*   It Creates a passworded MDB to stop other accessing your errors, you then can make
'*   a frontend to read your errors.
'*   Table Name : ErrList
'*   Field Name : ErrDate, ErrDes, ErrNum, ErrNotes, ErrUser       '*   'Usage
'*   Error_Handler:
'*   Select Case Error_Handler_Doc("Name.mdb", Now, 123, "Description", "Notes")
'*   Case "True"
'*   Case "False"
'*   End Select
'*   Load in "References" the "Microsoft DAO 3.51 Object Library"
  Dim NewDB As Database
  Dim ExistDB As Database
  Dim ExistRS As Recordset
Public Function Error_Handler_Doc(ByVal ErrMDB As String, ErrDate As Date, ErrNum As Long, ErrDes As String, ErrNote As String, Optional ErrUser As String) As Boolean
Select Case Error_Handler_MDB(ErrMDB)
  Case "False"
    If Error_Handler_Create(ErrMDB, "!@#$") = False Then
      Error_Handler_Doc = False
      Exit Function
    End If
End Select
  Set ExistDB = OpenDatabase("C:\Program Files\Common Files\Walker Brothers\ErrorHandler\" & ErrMDB, False, False, ";pwd=!@#$")
  Set ExistRS = ExistDB.OpenRecordset("ErrList", dbOpenDynaset)
    ExistRS.AddNew
    ExistRS.Fields!ErrNum = ErrNum & ""
    ExistRS.Fields!ErrDate = ErrDate & ""
    ExistRS.Fields!ErrDes = ErrDes & ""
    ExistRS.Fields!ErrNote = ErrNote & ""
    ExistRS.Fields!ErrUser = ErrUser & ""
    ExistRS.Update
  ExistRS.Close
  ExistDB.Close
  Set ExistRS = Nothing
  Set ExistDB = Nothing
  Error_Handler_Doc = True
End Function
Public Function Error_Handler_MDB(ByVal ErrMDB As String) As Boolean
  On Error Resume Next
  Open "C:\Program Files\Common Files\Walker Brothers\ErrorHandler\" & ErrMDB For Input As #1
    If Err Then
      Error_Handler_MDB = False
      Exit Function
    End If
  Close #1
  Error_Handler_MDB = True
End Function
Public Function Error_Handler_Create(ByVal ErrMDB As String, ByVal ErrMDBPassword As String) As Boolean
  Error_Handler_Create = False
  If CreateNewDirectory("C:\Program Files\Common Files\Walker Brothers\ErrorHandler") = False Then
    Exit Function
  End If
  On Error GoTo Err_Handler
  If ErrMDBPassword <> "" Then
    Set NewDB = Workspaces(0).CreateDatabase("C:\Program Files\Common Files\Walker Brothers\ErrorHandler\" & ErrMDB, dbLangGeneral & ";pwd=" & ErrMDBPassword)
  Else
    Set NewDB = Workspaces(0).CreateDatabase("C:\Program Files\Common Files\Walker Brothers\ErrorHandler\" & ErrMDB, dbLangGeneral)
  End If
  'Now call the functions for each table
  Dim b As Boolean
  b = Error_Handler_Err_List
  If b = False Then
    Error_Handler_Create = False
    NewDB.Close
    Set NewDB = Nothing
    Exit Function
  End If
  Error_Handler_Create = True
  SetAttr "C:\Program Files\Common Files\Walker Brothers\ErrorHandler\" & ErrMDB, vbHidden
  Exit Function
Err_Handler:
    If Err.Number <> 0 Then
        Error_Handler_Create = False
        NewDB.Close
        Set NewDB = Nothing
        Exit Function
    End If
End Function
Public Function Error_Handler_Err_List() As Boolean
  Dim TempTDef As TableDef
  Dim TempField As Field
  Dim TempIdx As Index
  Error_Handler_Err_List = False
  On Error GoTo Err_Handler
  Set TempTDef = NewDB.CreateTableDef("ErrList")
    Set TempField = TempTDef.CreateField("ErrDate", 8)
      TempField.Attributes = 1
      TempField.Required = False
      TempField.OrdinalPosition = 0
    TempTDef.Fields.Append TempField
    TempTDef.Fields.Refresh
    Set TempField = TempTDef.CreateField("ErrNum", 4)
      TempField.Attributes = 1
      TempField.Required = False
      TempField.OrdinalPosition = 1
    TempTDef.Fields.Append TempField
    TempTDef.Fields.Refresh
    Set TempField = TempTDef.CreateField("ErrDes", 12)
      TempField.Attributes = 2
      TempField.Required = False
      TempField.OrdinalPosition = 2
      TempField.AllowZeroLength = False
    TempTDef.Fields.Append TempField
    TempTDef.Fields.Refresh
    Set TempField = TempTDef.CreateField("ErrNote", 12)
      TempField.Attributes = 2
      TempField.Required = False
      TempField.OrdinalPosition = 3
      TempField.AllowZeroLength = False
    TempTDef.Fields.Append TempField
    TempTDef.Fields.Refresh
    Set TempField = TempTDef.CreateField("ErrUser", 10)
      TempField.Attributes = 2
      TempField.Required = False
      TempField.OrdinalPosition = 4
      TempField.Size = 50
      TempField.AllowZeroLength = True
    TempTDef.Fields.Append TempField
    TempTDef.Fields.Refresh
  NewDB.TableDefs.Append TempTDef
  NewDB.TableDefs.Refresh
  'Done, Close the objects
    Set TempTDef = Nothing
    Set TempField = Nothing
    Set TempIdx = Nothing
  Error_Handler_Err_List = True
  Exit Function
Err_Handler:
    If Err.Number <> 0 Then
    Set TempTDef = Nothing
    Set TempField = Nothing
    Set TempIdx = Nothing
    Error_Handler_Err_List = False
    Exit Function
    End If
End Function
Public Function CreateNewDirectory(ByVal NewDirectory As String) As Boolean
  Dim sDirTest As String
  Dim SecAttrib As SECURITY_ATTRIBUTES
  Dim bSuccess As Boolean
  Dim sPath As String
  Dim iCounter As Integer
  Dim sTempDir As String
  Dim iFlag As Integer
  On Error GoTo ErrorCreate
    iFlag = 0
    sPath = NewDirectory
    If Right(sPath, Len(sPath)) <> "\" Then
      sPath = sPath & "\"
    End If
    iCounter = 1
    Do Until InStr(iCounter, sPath, "\") = 0
      iCounter = InStr(iCounter, sPath, "\")
      sTempDir = Left(sPath, iCounter)
      sDirTest = Dir(sTempDir)
      iCounter = iCounter + 1
      'create directory
      SecAttrib.lpSecurityDescriptor = &O0
      SecAttrib.bInheritHandle = False
      SecAttrib.nLength = Len(SecAttrib)
      bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop
  CreateNewDirectory = True
  Exit Function
ErrorCreate:
  CreateNewDirectory = False
  Resume 0
End Function
'  'Usage
'  Select Case Error_Handler_Doc("Name.mdb", Now, 123, "Description", "Notes")
'    Case "True"
'    Case "False"
'  End Select
```

