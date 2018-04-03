Attribute VB_Name = "NamesFilesIO"

Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const WM_CLOSE = &H10


Dim FClientID As String

Public Function ReadINI(Section, KeyName, FileName As String) As String
    Dim sRet As String
    sRet = String(510, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Public Sub OpenRawFile(FileName As String, ClientID As String)

    FClientID = ClientID
    Open FileName For Input Access Read As #1

End Sub
Public Sub OpenErrorFile()
Dim FileName As String, LogDateTime As String
    LogDateTime = Format(Now, "yyyymmdd hhmmss")
    FileName = "ErrorList " & LogDateTime
    Open FileName For Input Access Read Lock Write As #1
End Sub
Public Sub OpenWordFile(FilenName As String, ClientID As String)

    FClientID = ClientID
   ' Open FileName For Input Access Read Lock Write As #1

End Sub
Public Sub OpenRawFileBinary(FileName As String, ClientID As String)

    FClientID = ClientID
    Open FileName For Binary Access Read As #1

End Sub

Public Sub OpenCSVFile(FileName As String)
    
    Open FileName For Output Access Write As #2

End Sub


Public Function ExportToTextFile(ByVal FName As String, ByVal OutPutFile As String, _
                            ByVal Sep As String, SelectionOnly As Boolean, Optional ColCount As Long) As Boolean

Dim WholeLine As String, FNum As Integer
Dim RowNdx As Long, ColNdx As Long, wsNameLen As Integer
Dim StartRow As Long, EndRow As Long, IStartedXL As Boolean
Dim StartCol As Long, EndCol As Long
Dim CellValue As String, xlApp As Excel.Application
Dim ws As Excel.Worksheet, wb As Excel.Workbook
Dim WinHandle As Long, PrintAll As Boolean
    On Error GoTo ErrorHandler
    Open OutPutFile For Output Access Write As #5
    KillProcess "EXCEL"
    IStartedXL = False
    'On Error Resume Next
    Set xlApp = New Excel.Application
    On Error GoTo 0
    IStartedXL = True
    Set wb = xlApp.Workbooks.Open(FileName:=FName)
    Set ws = wb.ActiveSheet
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    PrintAll = False
    
    For Each ws In wb.Worksheets
        If SelectionOnly = True Then
            With Selection
                StartRow = .Cells(1).row
                StartCol = .Cells(1).Column
                EndRow = .Cells(.Cells.Count).row
                EndCol = .Cells(.Cells.Count).Column
            End With
        Else
            ws.Activate
            With xlApp.ActiveSheet.UsedRange
                StartRow = .Cells(1).row
                StartCol = .Cells(1).Column
                EndRow = .Cells.Find(What:="*", _
                    SearchDirection:=xlPrevious, _
                    SearchOrder:=xlByRows).row
                If ColCount = 0 Then
                    EndCol = .Cells.Find(What:="*", _
                        SearchDirection:=xlPrevious, _
                        SearchOrder:=xlByColumns).Column
                Else
                    EndCol = ColCount
                End If
            End With
        End If
        wsNameLen = 0
        For RowNdx = StartRow To EndRow
            If InStr(OutPutFile, "La Amistad") > 0 Then
                WholeLine = ""
            ElseIf InStr(OutPutFile, "Aldrich") > 0 Then
                WholeLine = ""
                PrintAll = True
            ElseIf InStr(OutPutFile, "BradCard") > 0 Then
                WholeLine = ""
            Else
                WholeLine = ws.Name & Sep
                wsNameLen = Len(WholeLine)
            End If
            For ColNdx = StartCol To EndCol
                If xlApp.Cells(RowNdx, ColNdx).Value = "" Then
                    CellValue = Chr(34) & Chr(34)
                Else
                   CellValue = xlApp.Cells(RowNdx, ColNdx).Value
                End If
                WholeLine = WholeLine & CellValue & Sep
            Next ColNdx
            WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))
            If InStr(OutPutFile, "La Amistad") > 0 Then
                If InStr(WholeLine, "=") > 0 Then GoTo ErrorHandler
                If InStr(WholeLine, "M.R.#") = 0 Then
                    Print #5, WholeLine
                End If
            Else
                If Not PrintAll Then
                    If Mid(WholeLine, wsNameLen + 1, 5) <> "Acc #" Then
                        If Mid(WholeLine, wsNameLen + 1, 5) <> Chr(34) & Chr(34) & Sep & Chr(34) & Chr(34) Then
                            Print #5, WholeLine
                        End If
                    End If
                Else
                    Print #5, WholeLine
                End If
            End If
        Next RowNdx
    Next ws
    xlApp.DisplayAlerts = False
    
    
ErrorHandlerExit:
    wb.Close False, FName
    ExportToTextFile = True
    On Error Resume Next
    Close #5
    xlApp.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Application.ScreenUpdating = True
    KillProcess "EXCEL"
    'Kill FName
    Exit Function
ErrorHandler:
    ExportToTextFile = False
    Resume ErrorHandlerExit

End Function

Public Sub KillProcess(ByVal processName As String)
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename
          On Error GoTo ErrHandler
          Set oWMI = GetObject("winmgmts:")
          Set oServices = oWMI.InstancesOf("win32_process")

          For Each oService In oServices
                 servicename = LCase(Trim(CStr(oService.Name) & ""))
                 If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
                    ret = oService.Terminate
                    Exit For
                 End If
          Next
          Set oServices = Nothing
          Set oWMI = Nothing
 
ErrHandler:
  Err.Clear
End Sub


Public Function EOFRawFile()

    EOFRawFile = EOF(1)
    
End Function

Public Sub CloseRawFile()

    Close #1

End Sub
Public Sub MoveFile(InputFileName As String, BackupFileName As String)
    Dim BackupPath As String, BackupFile As String, i As Integer
    
    For i = Len(BackupFileName) To 1 Step -1
        If Mid(BackupFileName, i, 1) = "\" Then
            BackupPath = Left(BackupFileName, i - 1)
            BackupFile = Mid(BackupFileName, i + 1)
            Exit For
        End If
    Next i
    If InStr(1, BackupPath, "{DATE}", vbTextCompare) <> 0 Then
        BackupPath = Replace(BackupPath, "{DATE}", Format(Date, "YYYY-MM-DD"), 1, -1, vbTextCompare)
    End If
    If Not DirExists(BackupPath) Then
        MkDir (BackupPath)
    End If
    'BackupPath = "\\Umb-loader\Imports\"
    Name InputFileName As BackupPath + "\" + FileDtTm + BackupFile
    
End Sub

Public Sub CopyFile(InputFileName As String, BackupFileName As String)
    Dim i As Integer, objFSO As New FileSystemObject
    Dim BackupPath As String, BackupFile As String
    Dim FileDtTm As String
    
    FileDtTm = Format(Now, "yyyy-mm-dd hhmmss")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If InStr(1, BackupFileName, "{DATE}", vbTextCompare) <> 0 Then
        BackupFileName = Replace(BackupFileName, "{DATE}", Format(Date, "YYYY-MM-DD"), 1, -1, vbTextCompare)
    End If
    BackupPath = Left(BackupFileName, InStrRev(BackupFileName, "\"))
    BackupFile = Mid(BackupFileName, InStrRev(BackupFileName, "\") + 1)
    If Not objFSO.FolderExists(BackupPath) Then
        MkDir (BackupPath)
    End If
    objFSO.CopyFile InputFileName, BackupPath + FileDtTm + BackupFile
    
    Set objFSO = Nothing
End Sub

Public Sub CloseCSVFile()

    Close #2

End Sub
Public Sub CloseCSVFile2()

    Close #3

End Sub

Public Sub ReadRawFileLine(InputString As String)
    
    Line Input #1, InputString

End Sub

Public Sub ReadRawFileBlock(InputString As String, BlockSize As Integer)
    
    InputString = Input(BlockSize, #1)

End Sub

Public Sub ReadRawFileBinary(InputString As String, TermChars As String)
    Dim mychar As String, i As Integer
    
    InputString = ""
    Do While Not EOF(1)
        mychar = Input(1, #1)
        For i = 1 To Len(TermChars)
            If mychar = Mid(TermChars, i, 1) Then
                Exit Do
            End If
            InputString = InputString + mychar
        Next i
    Loop

End Sub

Public Sub ReadRawFields1(Field1)
    
On Error GoTo HandleReadError
    
    Input #1, Field1
    Exit Sub
    
HandleReadError:

    If Err.Number <> 62 Then
        Err.Raise (Err.Number)
    End If
    

End Sub

Public Sub ReadRawFields5(Field1 As String, Field2 As String, Field3 As String, _
                          Field4 As String, Field5 As String)
    
On Error GoTo HandleReadError
    
    Input #1, Field1, Field2, Field3, Field4, Field5
    Exit Sub
    
HandleReadError:

    If Err.Number <> 62 Then
        Err.Raise (Err.Number)
    End If
    

End Sub

Public Sub ReadRawFields6(Field1 As String, Field2 As String, Field3 As String, _
                          Field4 As String, Field5 As String, Field6 As String)
    
On Error GoTo HandleReadError
    
    Input #1, Field1, Field2, Field3, Field4, Field5, Field6
    Exit Sub
    
HandleReadError:

    If Err.Number <> 62 Then
        Err.Raise (Err.Number)
    End If
    

End Sub

Public Sub ReadRawFields7(Field1 As String, Field2 As String, Field3 As String, Field4 As String, Field5 As String, Field6 As String, Field7 As String)
    
On Error GoTo HandleReadError
    
    Input #1, Field1, Field2, Field3, Field4, Field5, Field6, Field7
    Exit Sub
    
HandleReadError:

    If Err.Number <> 62 Then
        Err.Raise (Err.Number)
    End If
    

End Sub

Public Sub ReadRawFields8(Field1 As String, Field2 As String, Field3 As String, Field4 As String, Field5 As String, Field6 As String, Field7 As String, Field8 As String)
    
On Error GoTo HandleReadError
    
    Input #1, Field1, Field2, Field3, Field4, Field5, Field6, Field7, Field8
    Exit Sub
    
HandleReadError:

    If Err.Number <> 62 Then
        Err.Raise (Err.Number)
    End If
    

End Sub

Public Sub WriteCSVFile(FieldsList As String)
    Dim OutputString  As String
    
    OutputString = FClientID + ","
    OutputString = OutputString + GetDateNumber(FieldsList, "PractDictatorID") + ","
    OutputString = OutputString + GetString(FieldsList, "ChartNo") + ","
    OutputString = OutputString + GetDateNumber(FieldsList, "DateOfService") + ","
    OutputString = OutputString + GetString(FieldsList, "PtNamePrefix") + ","
    OutputString = OutputString + GetString(FieldsList, "PtFirstName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtMiddleName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtLastName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtNameSuffix") + ","
    OutputString = OutputString + GetString(FieldsList, "PtFullName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtAddress1") + ","
    OutputString = OutputString + GetString(FieldsList, "PtAddress2") + ","
    OutputString = OutputString + GetString(FieldsList, "PtAddress3") + ","
    OutputString = OutputString + GetDateNumber(FieldsList, "PtDOB") + ","
    OutputString = OutputString + GetString(FieldsList, "RdNamePrefix") + ","
    OutputString = OutputString + GetString(FieldsList, "RdFirstName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdMiddleName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdLastName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdNameSuffix") + ","
    OutputString = OutputString + GetString(FieldsList, "RdFullName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdAddress1") + ","
    OutputString = OutputString + GetString(FieldsList, "RdAddress2") + ","
    OutputString = OutputString + GetString(FieldsList, "RdAddress3") + ","
    OutputString = OutputString + GetString(FieldsList, "RdFaxNumber") + ","
    OutputString = OutputString + GetString(FieldsList, "RdEmailAddress") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional1") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional2") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional3") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional4") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional5") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional6") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional7") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional8") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional9") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional10") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional11") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional12")
    Debug.Print OutputString
    Print #2, OutputString
    
End Sub
Public Sub WritePSIFiles(FieldsList As String, ByVal n As Integer, ByVal ClientID As String)
    Dim OutputString  As String
    
    OutputString = ClientID + ","
    OutputString = OutputString + GetDateNumber(FieldsList, "PractDictatorID") + ","
    OutputString = OutputString + GetString(FieldsList, "ChartNo") + ","
    OutputString = OutputString + GetDateNumber(FieldsList, "DateOfService") + ","
    OutputString = OutputString + GetString(FieldsList, "PtNamePrefix") + ","
    OutputString = OutputString + GetString(FieldsList, "PtFirstName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtMiddleName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtLastName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtNameSuffix") + ","
    OutputString = OutputString + GetString(FieldsList, "PtFullName") + ","
    OutputString = OutputString + GetString(FieldsList, "PtAddress1") + ","
    OutputString = OutputString + GetString(FieldsList, "PtAddress2") + ","
    OutputString = OutputString + GetString(FieldsList, "PtAddress3") + ","
    OutputString = OutputString + GetDateNumber(FieldsList, "PtDOB") + ","
    OutputString = OutputString + GetString(FieldsList, "RdNamePrefix") + ","
    OutputString = OutputString + GetString(FieldsList, "RdFirstName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdMiddleName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdLastName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdNameSuffix") + ","
    OutputString = OutputString + GetString(FieldsList, "RdFullName") + ","
    OutputString = OutputString + GetString(FieldsList, "RdAddress1") + ","
    OutputString = OutputString + GetString(FieldsList, "RdAddress2") + ","
    OutputString = OutputString + GetString(FieldsList, "RdAddress3") + ","
    OutputString = OutputString + GetString(FieldsList, "RdFaxNumber") + ","
    OutputString = OutputString + GetString(FieldsList, "RdEmailAddress") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional1") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional2") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional3") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional4") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional5") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional6") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional7") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional8") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional9") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional10") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional11") + ","
    OutputString = OutputString + GetString(FieldsList, "Optional12")
    Debug.Print OutputString
    Select Case n
        Case 2
          Print #2, OutputString
        Case 3
          Print #3, OutputString
    End Select
    
End Sub
Private Function GetString(FieldsList As String, KeyName As String)
    Dim Pos As Integer, pos2 As Integer
    
    Pos = InStr(FieldsList, KeyName)
    If Pos <> 0 Then
        Pos = Pos + Len(KeyName) + 1
        pos2 = InStr(Pos, FieldsList, "|")
        GetString = """" + Mid(FieldsList, Pos, pos2 - Pos) + """"
    Else
      GetString = """"""
    End If

End Function

Private Function GetDateNumber(FieldsList As String, KeyName As String)
    Dim Pos As Integer, pos2 As Integer
    
    Pos = InStr(FieldsList, KeyName)
    If Pos <> 0 Then
        Pos = Pos + Len(KeyName) + 1
        pos2 = InStr(Pos, FieldsList, "|")
        GetDateNumber = Mid(FieldsList, Pos, pos2 - Pos)
    Else
      GetDateNumber = ""
    End If

End Function

Public Function DirExists(Directory As String) As Boolean
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    If Dir(Directory, vbDirectory) = "." Then
        DirExists = True
    Else
        DirExists = False
    End If

End Function

Public Function FileExists(FileName As String) As Boolean
    If FileName = "" Then
        FileExists = False
    ElseIf Dir(FileName) <> "" Then
        FileExists = True
    Else
    If Err.Number <> "52" Then
        FileExists = False
    End If
    End If
    
End Function

Public Sub GetDictatorIDs(ClientID As String, NameType As String, DictArray() As String)
    Dim mdbFile As String
    Dim DBConnection As Database, rs As Recordset
    Dim SQL As String
    Dim i As Integer, NameRec As String

    mdbFile = ReadINI("FILE LOCATIONS", "Address Book", App.Path + "\ProcessNames.ini")

    SQL = "SELECT PractDictatorId, LastName, FirstName, MiddleName, Initials " + _
          "FROM DictatorsRecipients " + _
          "WHERE PracticeId = " + ClientID + " " + _
          "AND PractDictatorId IS NOT NULL"
    Set DBConnection = OpenDatabase(mdbFile)
    Set rs = DBConnection.OpenRecordset(SQL)
    
    i = -1
    While Not rs.EOF
        i = i + 1
        ReDim Preserve DictArray(i)
        If NameType = "L" Then
            DictArray(i) = UCase(rs.Fields!LastName)
        ElseIf NameType = "LFM" Then
            DictArray(i) = UCase(Trim(rs.Fields!LastName + rs.Fields!FirstName + rs.Fields!MiddleName))
        ElseIf NameType = "LF" Then
            DictArray(i) = UCase(Trim(rs.Fields!LastName + rs.Fields!FirstName))
        ElseIf NameType = "FML" Then
            If Not IsNull(rs.Fields!MiddleName) Then
                DictArray(i) = UCase(Trim(rs.Fields!FirstName + " " + rs.Fields!MiddleName + " " + rs.Fields!LastName))
            Else
                DictArray(i) = UCase(Trim(rs.Fields!FirstName + " " + rs.Fields!LastName))
            End If
        ElseIf NameType = "I" Then
            DictArray(i) = UCase(rs.Fields!Initials)
        End If
        DictArray(i) = DictArray(i) + "=" + CStr(rs.Fields!PractDictatorId)
        rs.MoveNext
    Wend
    Set rs = Nothing
    Set DBConnection = Nothing
    
    If i = -1 Then
      
      i = 1
      ReDim Preserve DictArray(i)
      DictArray(i) = ""
    End If
    
End Sub


