﻿Option Strict On

Imports System.IO
Imports Microsoft.VisualBasic.FileIO


Module ProcessNamesBHC
    Public Function ParseBHC(ByVal clientName As String, ByVal inputDir As String, ByVal inputFile As String, ByVal outputDir As String, _
                            ByVal backupDir As String) As Integer
        Dim InputString, FullName, ChartNo, DateOfBirth, outputFile, outputPath, errorPath As String
        Dim FirstName, MiddleName, LastName, Sex, DateOfService, DoctorID As String
        Dim ProcessState As Integer
        Dim StringExceptions() As String = {"WOUNDCARE", "OUT", "OUT OF TOW", "OUT OF TOWN", "MEETING", "ADD", "LUNCH", "BOARD", "ASSIST", "DENTIST"}
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

            'Set ProcessState to 1
            ProcessState = 1

            Using r As StreamReader = New StreamReader(inputDir & "\" & inputFile)
                Try
                    InputString = r.ReadLine
                    Do
                        blnIgnore = False
                        Select Case ProcessState
                            Case 1
                                'Parse out first demographic line with name information
                                If InputString <> "" Then
                                    LastName = Trim(Left(InputString, 16))
                                    FirstName = Trim(Mid(InputString, 17, 14))
                                    MiddleName = Trim(Mid(InputString, 31, 1))
                                    FullName = LastName & ", " & FirstName & " " & MiddleName
                                    If StringExceptions.Contains(LastName) Then
                                        blnIgnore = True
                                    End If
                                    ProcessState = 2
                                End If
                                'read next line
                                InputString = r.ReadLine
                            Case 2
                                'Parse out MRN, DateOfBirth and Sex
                                If blnIgnore Then
                                    ChartNo = "IGNORE"
                                Else
                                    ChartNo = Trim(Left(InputString, 10))
                                End If
                                DateOfBirth = Mid(InputString, 12, 2) & "/" & _
                                              Mid(InputString, 14, 2) & "/" & _
                                              Mid(InputString, 16, 4)
                                If Not IsNumeric(Left(DateOfBirth, 1)) Then DateOfBirth = "IGNORE"
                                Sex = Mid(InputString, 21, 1)
                                If Sex <> "M" And Sex <> "F" Then
                                    ChartNo = "IGNORE"
                                End If
                                ProcessState = 3
                                'read next line
                                InputString = r.ReadLine
                            Case 3
                                'Parse out DateOfService and DoctorID; write parsed info to output file
                                If Mid(InputString, 3, 1) = "/" And Mid(InputString, 6, 1) = "/" Then
                                    DateOfService = Right("0" + Trim(Left(InputString, 6) + "20" + _
                                                    Mid(InputString, 7, 2)), 10)
                                    DoctorID = Trim(Mid(InputString, 19, 5))

                                    If ChartNo <> "IGNORE" Then
                                        If DateOfBirth <> "IGNORE" Then
                                            WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                                                    "PtDOB=" + DateOfBirth + "|" + _
                                                    "PtFullName=" + FullName + "|" + _
                                                    "Optional1=" + Sex + "|" + _
                                                    "DateOfService=" + DateOfService + "|")
                                        Else
                                            WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                                                    "PtFullName=" + FullName + "|" + _
                                                    "Optional1=" + Sex + "|" + _
                                                    "DateOfService=" + DateOfService + "|")
                                        End If
                                        ProcessState = 1
                                    Else
                                        ProcessState = 3
                                    End If
                                End If

                                'read next line
                                InputString = r.ReadLine
                        End Select
                    Loop Until InputString Is Nothing
                Catch ex As Exception
                    If ErrMsg <> "" Then
                        ErrMsg = ErrMsg & "; " & ex.Message
                    Else
                        ErrMsg = ex.Message
                    End If
                Finally
                    r.Close()
                End Try
            End Using


        Catch ex As Exception
            ParseBHC = -1
            If ErrMsg <> "" Then
                ErrMsg = ErrMsg & "; " & ex.Message
            Else
                ErrMsg = ex.Message
            End If

            File.Move(inputDir & "\" & inputFile, errorPath & "\" & inputFile)
            Exit Function
        End Try
        ParseBHC = 0    'Successful completion
        File.Delete(inputDir & "\" & inputFile)
        ' CloseIntrascriptCSVFile(outputPath)

    End Function
End Module
