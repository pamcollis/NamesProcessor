Imports System.IO

Module ProcessNames_LCH
    Public Function ParseLCH(ByVal clientName As String, ByVal DoCase As String, ByVal inputDir As String, ByVal inputFile As String, ByVal outputDir As String, _
                            ByVal backupDir As String) As Integer


        Dim InputString, InputFilePath, FullName, ChartNo, DateOfBirth, outputFile, outputPath, errorPath As String
        Dim Dictator, inputFilename, HISBackupDir As String
        Dim streetAddress, homeCity, homeState, homeZipcode, FullAddress As String
        Dim FirstName As String, MiddleName As String, LastName As String
        Dim Prefix As String, Suffix As String
        Dim DateOfService As String
        Dim PatientInfo() As String
        Dim linecounter As Integer
        Dim FileDtTm As String = Format(Now, "yyyy-MM-dd hh-mm-ss")
        DateOfBirth = "" : ChartNo = "" : FullName = "" : outputPath = ""
        HISBackupDir = "C:\Intrascript\HIS\Lerner Cohen Healthcare\HIS upload\Backup"

        errorPath = Application.StartupPath & "\Errors"
        If (Not Directory.Exists(errorPath)) Then System.IO.Directory.CreateDirectory(errorPath)


        Try
            inputFilename = Dir(inputDir & "\" & inputFile)   'gets the complete filename
            outputFile = inputFilename
            'Copy input file to backup folder
            File.Copy(inputDir & "\" & inputFilename, backupDir & "\" & FileDtTm & " " & inputFilename)

            'Create output file
            outputFile = FileDtTm & " " & outputFile
            outputPath = outputDir & "\" & outputFile


            linecounter = 1

            If HISBackupDir <> "" Then

                File.Copy(inputDir & "\" & inputFilename, HISBackupDir & "\" & FileDtTm & " " & inputFilename)
            Else
                Err.Raise(513, "CopyFile", "Backup Directory does not exist")
            End If

            OpenIntrascriptCSVFile(outputPath)

            InputFilePath = inputDir & "\" & inputFilename

            Using sr As StreamReader = New StreamReader(InputFilePath)
                Try
                    InputString = sr.ReadLine
                    'While (InputString <> Nothing)
                    If InputString = Nothing Then GoTo NextLine
                    ReDim Preserve PatientInfo(0)
                    PatientInfo = Split(InputString, ",")
                    ReDim Preserve PatientInfo(UBound(PatientInfo))
                    While (InStr(PatientInfo(0), "End") = 0)
                        If PatientInfo.Length < 5 Then GoTo NextLine
                        If InStr(PatientInfo(0), "Title") = 0 Then
                            FirstName = Trim(Replace(PatientInfo(1), Chr(34), ""))
                            If String.IsNullOrEmpty(FirstName) <> True Then
                                If IsNumeric(Right(FirstName, 1)) Then GoTo NextLine
                                If UCase(Right(FirstName, 4)) = "CELL" Then
                                    If IsNumeric(Left(Right(FirstName, 7), 1)) Then GoTo NextLine
                                End If
                                If FirstName = "Switzerland" Then GoTo NextLine
                                If FirstName = "PA" Then GoTo NextLine
                                If FirstName = "2008 till May 2009 we will be at" Then GoTo NextLine
                                If FirstName = "Oh" Then GoTo NextLine
                                If FirstName = "Apt. 2A" Then GoTo NextLine
                                If FirstName = "Valley" Then GoTo NextLine
                                If FirstName = "000.00 renewal" Then GoTo NextLine
                                If InStr(FirstName, "address") > 0 Then GoTo NextLine
                                If Len(FirstName) > 20 Then GoTo NextLine
                                MiddleName = Trim(Replace(PatientInfo(2), Chr(34), ""))
                                If Len(MiddleName) > 20 Then GoTo NextLine
                                If IsNumeric(Right(MiddleName, 1)) Then GoTo NextLine
                                LastName = Trim(Replace(PatientInfo(3), Chr(34), ""))
                                If IsNumeric(Right(LastName, 1)) Then GoTo NextLine
                                Suffix = Trim(Replace(PatientInfo(4), Chr(34), ""))

                                If Not FirstName = "" Then
                                    If Not MiddleName = "" Then
                                        If Not Suffix = "" Then
                                            FullName = LastName & " " & Suffix & ", " & FirstName & " " & MiddleName
                                        Else
                                            FullName = LastName & ", " & FirstName & " " & MiddleName
                                        End If
                                    Else
                                        If Not Suffix = "" Then
                                            FullName = LastName & " " & Suffix & ", " & FirstName
                                        Else
                                            FullName = LastName & ", " & FirstName
                                        End If
                                    End If
                                Else
                                    GoTo NextLine
                                End If

                                Dictator = Trim(Replace(PatientInfo(5), Chr(34), ""))
                                If Not Dictator = "" Then
                                    If InStr(Dictator, "Cohen") > 0 Then
                                        Dictator = "Louis Cohen, M.D."
                                    ElseIf InStr(Dictator, "Lerner") > 0 Then
                                        Dictator = "Brad Lerner, M.D."
                                    ElseIf InStr(Dictator, "Cocco") > 0 Then
                                        Dictator = "James R. Cocco, M.D."
                                    End If

                                End If

                                DateOfBirth = Trim(Replace(PatientInfo(10), Chr(34), ""))
                                If InStr(DateOfBirth, "-") > 0 Then
                                    DateOfBirth = Replace(DateOfBirth, "-", "/")
                                End If
                                If DateOfBirth = "0/0/00" Then DateOfBirth = ""
                                If InStr(DateOfBirth, "14yrs") > 0 Then
                                    DateOfBirth = Trim(Replace(DateOfBirth, "14yrs old", " "))
                                End If
                                If Not DateOfBirth = "" Then
                                    DateOfBirth = FormattedDOB(DateOfBirth)
                                End If
                                DateOfBirth = Replace(DateOfBirth, " ", "")

                                streetAddress = Trim(Replace(PatientInfo(6), Chr(34), ""))
                                homeCity = Trim(Replace(PatientInfo(7), Chr(34), ""))
                                homeState = Trim(Replace(PatientInfo(8), Chr(34), ""))
                                homeZipcode = Trim(Replace(PatientInfo(9), Chr(34), ""))
                                FullAddress = streetAddress + "," + homeCity + "," + homeState + " " + homeZipcode
                                ChartNo = "999"

                                If Not FirstName = "" Then

                                    WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                                                    "PtDOB=" + DateOfBirth + "|" + _
                                                    "PtFullName=" + FullName + "|" + _
                                                    "PtAddress=" + FullAddress + "|" + _
                                                    "RdFullName=" + Dictator + "|")
                                    'Clear variables
                                    streetAddress = "" : ChartNo = ""
                                    FirstName = "" : MiddleName = "" : LastName = ""
                                    Suffix = "" : DateOfBirth = ""
                                    Dictator = "" : homeCity = "" : homeState = ""
                                End If
                            End If
                        End If
                        Prefix = ""
                        FirstName = ""
                        MiddleName = ""
                        LastName = ""
                        Suffix = ""
                        Dictator = ""
                        DateOfService = ""
                        DateOfBirth = ""
                        streetAddress = ""
                        homeCity = ""
                        homeState = ""
                        homeZipcode = ""
                        FullAddress = ""
NextLine:

                        InputString = sr.ReadLine
                        ReDim Preserve PatientInfo(0)
                        PatientInfo = Split(InputString, ",")
                        ReDim Preserve PatientInfo(UBound(PatientInfo))
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
            ParseLCH = -1
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
        ParseLCH = 0    'Successful completion
    End Function


    Private Function FormattedDate(ByVal fDate As String) As String
        Dim Slash1 As Integer, Slash2 As Integer
        Dim month As String, day As String, year As String
        Dim currYear As String
        currYear = Left(Format(Now, "yyyy"), 2)
        If InStr(fDate, "-") > 0 Then
            fDate = Replace(fDate, "-", "/")
        End If
        If InStr(fDate, "/") > 0 Then
            Slash1 = InStr(fDate, "/")
            Slash2 = InStr(Slash1 + 1, fDate, "/")
            month = Left(fDate, Slash1 - 1)
            day = Mid(fDate, Slash1 + 1, (Slash2 - Slash1) - 1)
            year = Mid(fDate, Slash2 + 1)

            If Len(month) = 1 Then month = "0" + month
            If Len(day) = 1 Then day = "0" + day
            If Len(year) = 2 Then year = currYear + year
            FormattedDate = month + "/" + day + "/" + year
        Else
            FormattedDate = ""
        End If

    End Function
    Private Function FormattedDOB(ByVal fDate As String) As String
        Dim Slash1 As Integer, Slash2 As Integer
        Dim month As String, day As String, year As String
        Dim currYear As String
        currYear = "19"
        If InStr(fDate, "-") > 0 Then
            fDate = Replace(fDate, "-", "/")
        End If
        If InStr(fDate, "/") > 0 Then
            Slash1 = InStr(fDate, "/")
            Slash2 = InStr(Slash1 + 1, fDate, "/")
            month = Left(fDate, Slash1 - 1)
            day = Mid(fDate, Slash1 + 1, (Slash2 - Slash1) - 1)
            year = Mid(fDate, Slash2 + 1)

            If Len(month) = 1 Then month = "0" + month
            If Len(day) = 1 Then day = "0" + day
            If Len(year) = 2 Then year = currYear + year
            FormattedDOB = month + "/" + day + "/" + year
        Else
            FormattedDOB = ""
        End If

    End Function


End Module
