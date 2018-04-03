
Imports System.IO

Module ProcessNamesLungA
    Public Function ParseLungA(ByVal clientName As String, ByVal DoCase As String, ByVal inputDir As String, ByVal inputFile As String, ByVal outputDir As String, _
                            ByVal backupDir As String) As Integer
        Dim InputString, FullName, ChartNo, DateOfService, outputFile, outputPath, errorPath As String
        Dim FirstName, MiddleName, LastName, Prefix, Suffix, ApptDesc, PrevDOS, PrevChartNo, PrevDesc As String
        Dim ParsedName As Object
        Dim lineCounter As Integer
        Dim StringExceptions() As String = {"WOUNDCARE", "OUT", "OUT OF TOW", "OUT OF TOWN", "MEETING", "ADD", "LUNCH", "BOARD", "ASSIST", "DENTIST"}
        Dim StringDesc() As String = {"F/U", "CT CHEST", "PFT", "CHECK UP", "CT/PE", "PET/CT", "CXR", "CXL"}
        Dim blnNextPatient, blnDescChecked, blnSkippedName As Boolean
        Dim FileDtTm As String = Format(Now, "yyyy-MM-dd hh-mm-ss")
        ChartNo = "" : DateOfService = "" : FullName = "" : outputPath = "" : ApptDesc = "" : PrevDOS = "" : PrevChartNo = "" : PrevDesc = ""
        blnNextPatient = False : blnDescChecked = False : blnSkippedName = False
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

                        If InStr(InputString, "RP00002") > 0 Then
                            lineCounter = 1
                            Do
                                'read next line
                                InputString = sr.ReadLine
                                lineCounter += 1
                            Loop Until lineCounter = 10
                        ElseIf InStr(InputString, "Date Appointment Scheduled:") > 0 Then
                            Exit While
                        End If

                        DateOfService = Trim(Mid(InputString, 6, 10))
                        If Not IsDate(DateOfService) Then
                            If blnSkippedName = False Then
                                'Loop thru Desc array to see if desc found in InputString
                                For Each sDesc As String In StringDesc
                                    If InStr(InputString, sDesc) > 0 Then
                                        blnNextPatient = False
                                        ApptDesc += Trim(InputString)
                                    End If
                                Next
                                blnDescChecked = True
                                GoTo NextLine
                            End If
                        Else


                            ChartNo = Trim(Mid(InputString, 38, 7))
                            If Left(ChartNo, 4) = "9999" Then
                                blnSkippedName = True
                                GoTo NextLine
                            End If

                            If ChartNo = "026680" Then
                                blnSkippedName = True
                                GoTo NextLine 'DRUG REP LUNCH'
                            End If
                            PrevDesc = ApptDesc
                            If PrevChartNo <> "" Then
                                If ChartNo <> PrevChartNo Then
                                    WriteCSVIntrascriptFile(outputPath, "ChartNo=" + PrevChartNo + "|" + _
                                                            "PtFullName=" + FullName + "|" + _
                                                            "PatientLocation=" + PrevDesc + "|" + _
                                                            "ServiceDate=" + PrevDOS + "|")
                                    ApptDesc = ""
                                    blnDescChecked = False
                                    blnSkippedName = False
                                Else
                                    If DateOfService <> PrevDOS Then
                                        WriteCSVIntrascriptFile(outputPath, "ChartNo=" + PrevChartNo + "|" + _
                                                               "PtFullName=" + FullName + "|" + _
                                                               "PatientLocation=" + PrevDesc + "|" + _
                                                               "ServiceDate=" + PrevDOS + "|")
                                        ApptDesc = ""
                                        blnDescChecked = False
                                        blnSkippedName = False
                                    End If
                                End If


                            End If
                            PrevDOS = DateOfService
                            PrevChartNo = ChartNo

                            FullName = Trim(Mid(InputString, 50, 38))
                            If StringExceptions.Contains(FullName) Then
                                GoTo NextLine
                            End If

                            ParsedName = ParseFMLName(FullName, DoCase)
                            Prefix = ParsedName(0)
                            FirstName = ParsedName(1)
                            MiddleName = ParsedName(2)
                            LastName = ParsedName(3)
                            Suffix = ParsedName(4)
                            If Not MiddleName = "" Then
                                FullName = LastName + ", " + FirstName + " " + MiddleName
                            Else
                                FullName = LastName + ", " + FirstName
                            End If

                            'If blnDescChecked Then
WriteToFile:

                            '    WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                            '                             "PtFullName=" + FullName + "|" + _
                            '                             "PatientType=" + ApptDesc + "|" + _
                            '                             "ServiceDate=" + DateOfService + "|")
                            '    Else
                            '    WriteCSVIntrascriptFile(outputPath, "ChartNo=" + ChartNo + "|" + _
                            '                             "PtFullName=" + FullName + "|" + _
                            '                             "ServiceDate=" + DateOfService + "|")
                            'End If
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
            ParseLungA = -1
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
        ParseLungA = 0    'Successful completion
    End Function

End Module

