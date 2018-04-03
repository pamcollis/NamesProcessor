Attribute VB_Name = "MainModule"
Option Explicit
Public ClientID As String, FileDtTm As String
Dim CommandLine As String

Sub Main()
    Dim clClient As String
    Dim clInputFile As String, clOutputFile As String
    Dim i As Integer, AutoRun As String
    
    CommandLine = Trim(Command())
'    CommandLine = "/r " + _
'                  "/C=239642000" + _
'                  "/I=04232004-165539-MGNames_INPatient-H#042304-3at02-A-manateeglensnames-04232004-04232004-144613.TXT.txt" + _
'                  "/O=TestMGInNew.txt"
    
    CommandLine = "/A "

    If UCase(Left(CommandLine, 2)) = "/R" Then
        If InStr(UCase(CommandLine), "/C=") <> 0 Then
            clClient = Trim(Mid(CommandLine, InStr(UCase(CommandLine), "/C=") + 3, _
                            InStr(InStr(UCase(CommandLine), "/C=") + 3, CommandLine, "/") - (InStr(UCase(CommandLine), "/C=") + 3) - 1))
            ClientID = clClient
            If InStr(UCase(CommandLine), "/I=") <> 0 Then
                clInputFile = Trim(Mid(CommandLine, InStr(UCase(CommandLine), "/I=") + 3, _
                                InStr(InStr(UCase(CommandLine), "/I=") + 3, CommandLine, "/") - (InStr(UCase(CommandLine), "/I=") + 3) - 1))
            End If
            clOutputFile = ReadINI("C" + clClient, "OutputFile", App.Path + "\ProcessNames.ini")
            If InStr(UCase(CommandLine), "/O=") <> 0 Then
                clOutputFile = Trim(Mid(CommandLine, InStr(UCase(CommandLine), "/O=") + 3))
            End If
            ProcessNamesMain clClient, clInputFile, clOutputFile
        Else
            MsgBox "No client specified.", vbCritical
        End If
    ElseIf UCase(Left(CommandLine, 2)) = "/A" Then
        Load frmMain
        For i = frmMain.cbClients.ListCount - 1 To 0 Step -1
            AutoRun = ReadINI("C" + frmMain.lbClients.List(frmMain.cbClients.ItemData(i)), "AutoRun", App.Path + "\ProcessNames.ini")
            If UCase(AutoRun) <> "TRUE" Then
                frmMain.cbClients.RemoveItem i
            End If
        Next i
        If frmMain.cbClients.ListCount = 0 Then
            MsgBox "No clients are Auto Run.", vbCritical
        Else
            frmMain.Caption = "Names Processor (Auto Run)"
            frmMain.cbClients.Visible = False
            frmMain.Label1.Item(0).Visible = False
            frmMain.edtFileName.Visible = False
            frmMain.btnFind.Visible = False
            frmMain.Label2.Item(0).Visible = False
            frmMain.btnOK.Visible = False
            frmMain.lvLog.Visible = True
            frmMain.btnCancel.Left = CInt((frmMain.Width - frmMain.btnCancel.Width) / 2)
            frmMain.Timer1.Enabled = True
            frmMain.Show (1)
        End If
    Else
        frmMain.Show (1)
    End If
    End
    
End Sub

Function ProcessNamesMain(Client As String, InputFile As String, OutPutFile As String) As Integer

    Dim clClient As String, clDoCase As String
    Dim clInputDir As String, clInputFile As String
    Dim clOutputDir As String, clOutputFile As String
    Dim clBackupDir As String, clTempDir As String, clConvert As String, i As Integer
    Dim LogFile As Integer, ErrMsg As String, LogMessage As String
    Dim ListItem As Object

    clClient = Client
    FileDtTm = Format(Now, "yyyy-mm-dd hh-nn-ss ")
    If InputFile <> "" Then
        clInputFile = InputFile
    Else
        clInputFile = ReadINI("C" + clClient, "InputFile", App.Path + "\ProcessNames.ini")
    End If
    If OutPutFile <> "" Then
        clOutputFile = OutPutFile
    Else
        clOutputFile = ReadINI("C" + clClient, "OutputFile", App.Path + "\ProcessNames.ini")
    End If
    clConvert = UCase(ReadINI("C" + clClient, "Convert", App.Path + "\ProcessNames.ini"))
    clDoCase = UCase(ReadINI("C" + clClient, "DoCase", App.Path + "\ProcessNames.ini"))
    clInputDir = ReadINI("C" + clClient, "InputDir", App.Path + "\ProcessNames.ini")
    clOutputDir = ReadINI("C" + clClient, "OutputDir", App.Path + "\ProcessNames.ini")
    clBackupDir = ReadINI("C" + clClient, "BackupDir", App.Path + "\ProcessNames.ini")
    clTempDir = ReadINI("C" + clClient, "TempDir", App.Path + "\ProcessNames.ini")
    If clConvert <> "" Then
        If InStr(clInputFile, "\") <> 0 Then
            For i = Len(clInputFile) To 1 Step -1
                If Mid(clInputFile, i, 1) = "\" Then
                    clInputDir = Left(clInputFile, i - 1)
                    clInputFile = Mid(clInputFile, i + 1)
                    Exit For
                End If
            Next i
        End If
        If InStr(clOutputFile, "\") <> 0 Then
            For i = Len(clOutputFile) To 1 Step -1
                If Mid(clOutputFile, i, 1) = "\" Then
                    clOutputDir = Left(clOutputFile, i - 1)
                    clOutputFile = Mid(clOutputFile, i + 1)
                    Exit For
                End If
            Next i
        End If
        If clOutputFile = "" Then
            clOutputFile = clInputFile
        End If
        clOutputFile = FileDtTm + ClientID + " " + clOutputFile
        While Not IsNumeric(clClient)
            clClient = Left(clClient, Len(clClient) - 1)
        Wend
       Select Case UCase(clConvert)
            Case "PROCESSNAMESAIM"
                i = ProcessNamesAIMNew(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESAIMAPPT"
                i = ProcessNamesAIMAppt(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESALDRICH"
                i = ProcessNamesAldrich(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESBATEYEMR"
                i = ProcessNamesBateyEMR(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESBRADCARD"
                i = ProcessNamesBradCard(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESBRADCARDXLS"
                i = ProcessNamesBradCardxls(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESCANYONMANOR"
                i = ProcessNamesCanyonManor(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESCCE"
                i = ProcessNamesCCE(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESCFFAD"
                i = ProcessNamesCFFAD(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESCHILDS"
                i = ProcessNamesChilds(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESFAIRFAX"
                i = ProcessNamesFairfax(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESGEO"
                i = ProcessNamesGEO(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESGEOSOUTHDISCH"
                i = ProcessNamesGEOSouthDisch(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESGEOSOUTHROSTER"
                i = ProcessNamesGEOSouthRoster(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESHEALTHWISE"
                i = ProcessNamesHealthwise(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESHEARTLAND"
                i = ProcessNamesHeartland(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESHOLLYHILL"
                i = ProcessNamesHollyHill(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESHCA"
                i = ProcessNamesHCA(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESIMA"
                i = ProcessNamesIMA(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESLAAMISTAD"
                i = ProcessNamesLaAmistad(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESLAURELRIDGE"
                i = ProcessNamesLaurelRidge(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESLCH"
                i = ProcessNamesLCH(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESLUNG"
                i = ProcessNamesLung(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMCDOWELL"
                i = ProcessNamesMcDowell(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMCRH"
                i = ProcessNamesMCRH(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            
            Case "PROCESSNAMESMANATEEGLENS"
                i = ProcessNamesManateeGlens(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMANATEEGLENSLAMP"
                i = ProcessNamesManateeGlensLAMP(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMHC"
                i = ProcessNamesMHC(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMEDIC"
                i = ProcessNamesMedic(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMEDICLB"
                i = ProcessNamesMedicLB(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMEDICVHC"
                i = ProcessNamesMedicVHC(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMEDINA"
                i = ProcessNamesMedina(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMEDINAHL7"
                i = ProcessNamesMedinaHL7(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESMORROW"
                i = ProcessNamesMorrow(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESNORTHSIDE"
                i = ProcessNamesNorthSide(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESPEAK"
                i = ProcessNamesPeak(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESPEAKPSI"
                i = ProcessNamesPeakPSI(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESPIW"
                i = ProcessNamesPIW(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESPIWpdf"
                i = ProcessNamesPIWpdf(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESPSIHOSP"
                i = ProcessNamesPSIHosp(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESSANDYPINES"
                i = ProcessNamesSandyPines(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESSAVE"
                i = ProcessNamesSAVE(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESUPD"
                i = ProcessNamesUPD(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case "PROCESSNAMESUHS"
                i = ProcessNamesUHS(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir, clTempDir)
            Case "PROCESSNAMESVHC"
                i = ProcessNamesVHC(clClient, clDoCase, _
                                    clInputDir, clInputFile, _
                                    clOutputDir, clOutputFile, _
                                    clBackupDir)
            Case Else
                i = -1
                ErrMsg = "Unknown Convert '" + clConvert + "' in INI file."
        End Select
        If i <> 0 Then
            ErrMsg = Error(i)
        End If
    Else
        i = -2
        ErrMsg = "Client specified not in INI file."
    End If
    If i = 0 Then
        If UCase(Left(CommandLine, 2)) <> "/R" And _
           UCase(Left(CommandLine, 2)) <> "/A" Then
            MsgBox "Name file processing completed successfully", vbInformation
        End If
        LogMessage = "        " + FileDtTm + Chr(9) + clConvert + Chr(9) + _
                     clClient + Chr(9) + _
                     clInputDir + "\" + clInputFile
        Set ListItem = frmMain.lvLog.ListItems.Add(, , FileDtTm)
        ListItem.SubItems(1) = clInputFile
        ListItem.SubItems(2) = clClient
        ListItem.Checked = True
    ElseIf i = -255 Then
        If UCase(Left(CommandLine, 2)) <> "/R" And _
           UCase(Left(CommandLine, 2)) <> "/A" Then
            MsgBox "Name file processing could not run:" + Chr(10) + _
                        clConvert + Chr(10) + _
                        clClient + Chr(10) + _
                        clInputDir + "\" + clInputFile + Chr(10) + _
                        "Some required files are not available.", vbCritical
        End If
        LogMessage = ""
    Else
        If UCase(Left(CommandLine, 2)) <> "/R" And _
           UCase(Left(CommandLine, 2)) <> "/A" Then
            MsgBox "Name file processing failed:" + Chr(10) + _
                        clConvert + Chr(10) + _
                        clClient + Chr(10) + _
                        clInputDir + "\" + clInputFile + Chr(10) + _
                        ErrMsg, vbExclamation
        Else
            MoveFile clInputDir + "\" + clInputFile, App.Path + "\Errors\" + clInputFile
        End If
        LogMessage = "Error - " + FileDtTm + Chr(9) + clConvert + Chr(9) + _
                     clClient + Chr(9) + _
                     clInputDir + "\" + clInputFile + Chr(9) + _
                     ErrMsg
        Set ListItem = frmMain.lvLog.ListItems.Add(, , FileDtTm)
        ListItem.SubItems(1) = clInputFile
        ListItem.SubItems(2) = clClient
        ListItem.Checked = False
    End If
    
WriteLogFile:
On Error GoTo WriteLogError

    LogFile = FreeFile
    If Not DirExists(App.Path + "\Logs") Then
        MkDir (App.Path + "\Logs")
    End If
    If Not DirExists(App.Path + "\Logs\" + Format(Date, "yyyy-mm-dd")) Then
        MkDir (App.Path + "\Logs\" + Format(Date, "yyyy-mm-dd"))
    End If
    Open App.Path + "\Logs\" + Format(Date, "yyyy-mm-dd") + "\ProcessNamesLog.txt" _
         For Append Access Write Lock Write As LogFile
    If LogMessage <> "" Then
        Print #LogFile, LogMessage
    End If
    
ExitProcessNamesMain:
On Error GoTo 0

    Close #LogFile
    ProcessNamesMain = i
    Exit Function
    
WriteLogError:

    Sleep 2
    GoTo WriteLogFile
    
End Function

Sub Sleep(Seconds As Integer)
    Dim StartTime As Timer
    
    StartTime = Timer
    While Timer < StartTime + Seconds
        DoEvents
    Wend

End Sub



