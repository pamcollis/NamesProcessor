Imports Microsoft.VisualBasic
Imports System.Timers
Imports System.IO
Imports System.IO.FileSystemInfo
Imports System.IO.DirectoryInfo
Imports System.Diagnostics
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Text



Public Class fMain

    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        
        ProcessNames()

    End Sub
    Private Sub ProcessNames()
        Dim clClient, clDoCase, iniPath As String
        Dim clInputDir As String, clInputFile As String
        Dim clOutputDir As String
        Dim clBackupDir As String, clConvert As String
        Dim i, p As Integer
        Dim listItems(4) As String
        Dim itm As ListViewItem

        ErrMsg = ""
        i = 1 : p = 1
        iniPath = Application.StartupPath & "\ProcessNames.ini"
        Dim ini As New IniFile()
        Do Until ini.IniReadValue("C" & i, "Name", iniPath) = ""
            clClient = ini.IniReadValue("C" & i, "Name", iniPath)
            clInputDir = ini.IniReadValue("C" & i, "InputDir", iniPath)
            clOutputDir = ini.IniReadValue("C" & i, "OutputDir", iniPath)
            clInputFile = ini.IniReadValue("C" & i, "InputFile", iniPath)
            clBackupDir = ini.IniReadValue("C" & i, "BackupDir", iniPath)
            clConvert = ini.IniReadValue("C" & i, "Convert", iniPath)
            clDoCase = ini.IniReadValue("C" & i, "DoCase", iniPath)
            ErrMsg = ""
            If clConvert = "" Then
                p = -2
                ErrMsg = "Client specified not in INI file"
            Else
                'If System.IO.File.Exists(clInputDir & "\" & clInputFile) Then
                If Dir(clInputDir & "\" & clInputFile) <> "" Then
                    Select Case clConvert
                        Case "ProcessNamesBradCard"
                            p = ParseBradCard(clClient, clDoCase, clInputDir, clInputFile, clOutputDir, clBackupDir)
                        Case "ProcessNamesLung"
                            p = ParseLung(clClient, clInputDir, clInputFile, clOutputDir, clBackupDir)
                        Case "ProcessNamesLungA"
                            p = ParseLungA(clClient, clDoCase, clInputDir, clInputFile, clOutputDir, clBackupDir)
                        Case "ProcessNamesLCH"
                            p = ParseLCH(clClient, clDoCase, clInputDir, clInputFile, clOutputDir, clBackupDir)
                        Case "ProcessNamesUND"
                            p = ParseUND(clClient, clDoCase, clInputDir, clInputFile, clOutputDir, clBackupDir)
                        Case Else
                            p = -1
                            ErrMsg = "Unknown Convert, " & clConvert & ", in INI file"
                    End Select
                End If

                i = i + 1
            End If
            If p < 0 Then ErrMsg = ErrorToString(p)
            If p = 0 Then
                'List as successfully parsed
                listItems(0) = clClient
                listItems(1) = Format(Now, "yyyy-MM-dd hh-mm-ss").ToString
                listItems(2) = "Successful"
                listItems(3) = ""
                itm = New ListViewItem(listItems)
                lvProcessedNames.Items.Add(itm)
                p = 1
            ElseIf p = 1 Then
                ' do nothing
            Else
                'List as Error with error message
                listItems(0) = clClient
                listItems(1) = Format(Now, "yyyy-MM-dd hh-mm-ss").ToString
                listItems(2) = "Error"
                listItems(3) = ErrMsg
                itm = New ListViewItem(listItems)
                lvProcessedNames.Items.Add(itm)
            End If
        Loop
        

    End Sub

End Class