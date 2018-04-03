Option Strict On

Imports System.IO

Module ProcessNamesLung
    Public Function ParseLung(ByVal clientName As String, ByVal inputDir As String, ByVal inputFile As String, ByVal outputDir As String, _
                            ByVal backupDir As String) As Integer
        Dim InputString, FullName, ChartNo, DateOfBirth, outputFile, outputPath, errorPath As String
        Dim FirstName, MiddleName, LastName, Sex As String
        Dim lineCounter As Integer
        Dim StringExceptions() As String = {"WOUNDCARE", "OUT", "OUT OF TOW", "OUT OF TOWN", "MEETING", "ADD", "LUNCH", "BOARD", "ASSIST", "DENTIST", "SMH"}
        Dim blnIgnore As Boolean = False
        Dim FileDtTm As String = Format(Now, "yyyy-MM-dd hh-mm-ss")
        DateOfBirth = "" : ChartNo = "" : Sex = "" : FullName = "" : outputPath = ""
        outputFile = inputFile
        errorPath = Application.StartupPath & "\Errors"
        If (Not Directory.Exists(errorPath)) Then System.IO.Directory.CreateDirectory(errorPath)

        Try
            'Copy input file to backup folder
            File.Copy(inputDir & "\" & inputFile, backupDir & "\" & FileDtTm & " " & clientName & " " & inputFile)

            'Create output file
            outputFile = FileDtTm & " " & clientName & " " & outputFile
            outputPath = outputDir & "\" & outputFile
            OpenIntrascriptCSVFile(outputPath)

            lineCounter = 1

            Using sr As StreamReader = New StreamReader(inputDir & "\" & inputFile)
                Try
                    InputString = sr.ReadLine
                    While (InputString <> Nothing)

                        If InStr(InputString, "RP00146") > 0 Then
                            lineCounter = 1
                            Do
                                'read next line
                                InputString = sr.ReadLine
                                lineCounter += 1
                            Loop Until lineCounter = 10
                        ElseIf InStr(InputString, "Total # of patients") > 0 Then
                            Exit While
                        End If

                        ChartNo = Trim(Left(InputString, 15))

                        LastName = Trim(Mid(InputString, 16, 16))
                        If (Left(LastName, 3) = "ZZZ") Then
                            GoTo NextLine
                        End If
                        If LastName = "" Then
                            Exit While
                        End If
                        FirstName = Trim(Mid(InputString, 32, 14))
                        MiddleName = Trim(Mid(InputString, 47, 1))
                        FullName = LastName & ", " & FirstName & " " & MiddleName
                        If StringExceptions.Contains(LastName) Then
                            blnIgnore = True
                        End If
                        DateOfBirth = Mid(InputString, 118, 10)
                        If Not IsDate(DateOfBirth) Then DateOfBirth = "IGNORE"

                        If DateOfBirth <> "IGNORE" Then
                            WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                                    "PtDOB=" + DateOfBirth + "|" + _
                                    "PtFullName=" + FullName + "|")
                        Else
                            WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                                    "PtFullName=" + FullName + "|")
                        End If
NextLine:

                        InputString = sr.ReadLine
                    End While
                Catch ex As Exception
                    If ErrMsg <> "" Then
                        ErrMsg = ErrMsg & "; " & ex.Message
                    Else
                        ErrMsg = ex.Message
                    End If
                Finally
                    sr.Close()
                End Try
            End Using


        Catch ex As Exception
            ParseLung = -1
            If ErrMsg <> "" Then
                ErrMsg = ErrMsg & "; " & ex.Message
            Else
                ErrMsg = ex.Message
            End If

            File.Move(inputDir & "\" & inputFile, errorPath & "\" & inputFile)
            Exit Function
        Finally
            File.Delete(inputDir & "\" & inputFile)
            ' CloseIntrascriptCSVFile(outputPath)
        End Try
        ParseLung = 0    'Successful completion
    End Function

End Module
