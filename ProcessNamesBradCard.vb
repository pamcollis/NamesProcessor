

Imports System.IO

Module ProcessNamesBradCard
    Public Function ParseBradCard(ByVal clientName As String, ByVal DoCase As String, ByVal InputPath As String, ByVal InputFileName As String, _
                                  ByVal OutputDir As String, ByVal BackupPath As String) As Integer
        Dim InputString As String
        Dim Dictator, OutputFileName, inputFile
        Dim FirstName, MiddleName, LastName As String
        Dim FullName, DOB, Prefix, Suffix As String
        Dim DateOfService, ChartNo As String
        Dim ApptTime, ApptType, OtherID As String
        Dim PatientInfo() As String, DataLine() As String
        Dim blnAtPatientInfo As Boolean
        Dim blnXlsToTxt As Boolean
        Dim ParsedName As Object
        Dim OutputPath, outputFile, FileDtTm, errorPath As String


        DateOfService = "" : OutputFileName = ""
        FileDtTm = Format(Now, "yyyy-MM-dd hh-mm-ss")
        errorPath = Application.StartupPath & "\Errors"
        If (Not Directory.Exists(errorPath)) Then System.IO.Directory.CreateDirectory(errorPath)


        OutputFileName = FileDtTm & " " & clientName & ".txt"
        inputFile = Dir(InputPath & "\" & InputFileName)   'gets the complete filename

            Try
                'Copy input file to backup folder
                File.Copy(InputPath & "\" & inputFile, BackupPath & "\" & FileDtTm & " " & clientName & " " & inputFile)

                blnXlsToTxt = ExportToTextFile(InputPath + "\" + inputFile, InputPath + "\" + "BradCardTemp.txt", Chr(9), False)
                If blnXlsToTxt Then
                    'Create output file

                    outputFile = OutputFileName
                    OutputPath = OutputDir & "\" & outputFile
                    OpenIntrascriptCSVFile(OutputPath)

                    ReDim Preserve PatientInfo(0)
                    blnAtPatientInfo = False
                    Using sr As StreamReader = New StreamReader(InputPath + "\" + "BradCardTemp.txt")
                        InputString = sr.ReadLine()
                        While (InputString <> Nothing)
                            If InStr(InputString, "Schedule: Date") > 0 Then
                                DataLine = Split(InputString, Chr(9))
                                DateOfService = Trim(DataLine(3))
                                InputString = sr.ReadLine
                                GoTo NextLine
                            End If
                            If InStr(InputString, "Schedule: Resource") > 0 Then
                                DataLine = Split(InputString, Chr(9))
                                Dictator = Trim(DataLine(3))
                                blnAtPatientInfo = True
                                InputString = sr.ReadLine
                                GoTo NextLine
                            End If
                            If blnAtPatientInfo Then
                                PatientInfo = Split(InputString, Chr(9))
                                If InStr(PatientInfo(0), "Count:") > 0 Then
                                    blnAtPatientInfo = False
                                    InputString = sr.ReadLine
                                    GoTo NextLine
                                End If
                                ReDim Preserve PatientInfo(UBound(PatientInfo))
                                ApptTime = Trim((PatientInfo(0)))
                                FullName = Trim(Left(PatientInfo(1), InStr(PatientInfo(1), "[") - 1))
                                OtherID = Trim(Mid(PatientInfo(1), InStr(PatientInfo(1), "[") + 1))
                                OtherID = Left(OtherID, Len(OtherID) - 1)
                                FullName = Trim(Replace(FullName, Chr(34), ""))
                                ParsedName = ParseLFMName(FullName, DoCase)
                                Prefix = ParsedName(0)
                                FirstName = ParsedName(1)
                                MiddleName = ParsedName(2)
                                LastName = ParsedName(3)
                                Suffix = ParsedName(4)
                                ChartNo = Trim(Replace(PatientInfo(4), Chr(34), ""))
                                If ChartNo = "" Then
                                    ChartNo = OtherID
                                    OtherID = ""
                                Else
                                    If InStr(1, ChartNo, "-") > 0 Then
                                        ChartNo = Left(ChartNo, InStr(1, ChartNo, "-") - 1)
                                    End If
                                    If Len(ChartNo) < 4 Then
                                        If Not IsNumeric(ChartNo) Then
                                            ChartNo = OtherID
                                            OtherID = ""
                                        End If
                                    End If
                                End If
                                DOB = PatientInfo(5)
                                ApptType = PatientInfo(6)
                                If ChartNo = "" Then ChartNo = OtherID
                                If OtherID = "" Then OtherID = ChartNo

                                WriteCSVIntrascriptFile(OutputPath, "ChartNo=" + OtherID + "|" + _
                                                     "AcctNo=" + ChartNo + "|" + _
                                                     "PtDOB=" + DOB + "|" + _
                                                     "PtFullName=" + FullName + "|" + _
                                                     "PatientLocation=" + ApptType + "|" + _
                                                     "ServiceDate=" + DateOfService + "|")

                                Prefix = ""
                                MiddleName = ""
                                LastName = ""
                                Suffix = ""
                                Dictator = ""
                                FirstName = ""
                            End If
                            InputString = sr.ReadLine

NextLine:
                        End While
                    End Using
                End If
            Catch ex As Exception
                ParseBradCard = -1
                If ErrMsg <> "" Then
                    ErrMsg = ErrMsg & "; " & ex.Message
                Else
                    ErrMsg = ex.Message
                End If
                File.Move(InputPath & "\" & inputFile, errorPath & "\" & inputFile)
                Exit Function
            End Try

            File.Delete(InputPath & "\" & inputFile)
            File.Delete(InputPath + "\" + "BradCardTemp.txt")

            ParseBradCard = 0  'Successful completion

    End Function
End Module
