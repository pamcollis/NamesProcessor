Imports System.IO

Module ProcessNamesUND
    Public Function ParseUND(ByVal clientName As String, ByVal DoCase As String, ByVal inputDir As String, ByVal inputFile As String, ByVal outputDir As String, _
                            ByVal backupDir As String) As Integer


        Dim InputString, FullName, ChartNo, DateOfBirth, outputFile, outputPath, errorPath As String
        Dim Provider, inputFilename, HISBackupDir As String
        Dim FirstName As String, MiddleInit As String, LastName As String
        Dim Prefix As String, Suffix As String
        Dim DateOfService
        Dim PatientInfo() As String
        Dim linecounter As Integer
        Dim FileDtTm As String = Format(Now, "yyyy-MM-dd hh-mm-ss")
        DateOfBirth = "" : ChartNo = "" : FullName = "" : outputPath = ""
        HISBackupDir = "C:\Intrascript\HIS\UND Family Medicine\HIS upload\Backup"

        errorPath = Application.StartupPath & "\Errors"
        If (Not Directory.Exists(errorPath)) Then System.IO.Directory.CreateDirectory(errorPath)

        inputFilename = Dir(inputDir & "\" & inputFile)   'gets the complete filename

        Try

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


            Using sr As StreamReader = New StreamReader(inputDir & "\" & inputFilename)
                Try
                    InputString = sr.ReadLine
                    InputString = Replace(InputString, Chr(34), "")
                    While (InputString <> Nothing)

                        ReDim Preserve PatientInfo(0)

                        PatientInfo = Split(InputString, ",")
                        ReDim Preserve PatientInfo(UBound(PatientInfo))
                        If Not PatientInfo(0) = "Patient Number" Then
                            Provider = PatientInfo(18) & ", " & PatientInfo(19)
                            Provider = Replace(Provider, Chr(34), "")
                            ChartNo = Replace(PatientInfo(11), Chr(34), "")
                            FirstName = Replace(PatientInfo(2), Chr(34), "")
                            MiddleInit = Replace(PatientInfo(3), Chr(34), "")
                            LastName = Replace(PatientInfo(1), Chr(34), "")
                            Suffix = Replace(PatientInfo(4), Chr(34), "")

                            If Not FirstName = "" Then
                                If Not MiddleInit = "" Then
                                    If Not Suffix = "" Then
                                        FullName = LastName & " " & Suffix & ", " & FirstName & " " & MiddleInit
                                    Else
                                        FullName = LastName & ", " & FirstName & " " & MiddleInit
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
                            FullName = FullName.ToUpper()
                            DateOfService = PatientInfo(14)    'csv value is appt date/time in m/d/yyyy format                           

                            DateOfBirth = PatientInfo(9)

                            If Not FirstName = "" Then

                                WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                                                "PtDOB=" + DateOfBirth + "|" + _
                                                "PtFullName=" + FullName + "|" + _
                                                "RdFullName=" + Provider + "|" + _
                                                "ServiceDate=" + DateOfService + "|")
                                'Clear variables
                                DateOfService = "" : ChartNo = ""
                                FirstName = "" : MiddleInit = "" : LastName = ""
                                Suffix = "" : DateOfBirth = ""
                                Provider = ""
                            End If
                        End If

                        Prefix = ""
                        FirstName = ""
                        MiddleInit = ""
                        LastName = ""
                        Suffix = ""
                        Provider = ""
                        DateOfService = ""
                        DateOfBirth = ""
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
            ParseUND = -1
            If ErrMsg <> "" Then
                ErrMsg = ErrMsg & "; " & ex.Message
            Else
                ErrMsg = ex.Message
            End If

            File.Move(inputDir & "\" & inputFile, errorPath & "\" & inputFile)
            Exit Function
        Finally
            'CloseIntrascriptCSVFile(inputDir & "\" & inputFilename)
            File.Delete(inputDir & "\" & inputFilename)
        End Try
        ParseUND = 0    'Successful completion
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
