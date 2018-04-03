
Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module NamesFilesIO
    Public Sub OpenRawFile(FileName As String)


    End Sub
    Public Sub OpenIntrascriptCSVFile(ByVal FileName As String)
        'Dim fs As FileStream = Nothing
        If (Not File.Exists(FileName)) Then
            'fs = File.Create(FileName)
            File.Create(FileName).Dispose()
        End If
    End Sub
    Public Sub CloseIntrascriptCSVFile(ByVal FileName As String)
        Dim fs As FileStream = Nothing
        If File.Exists(FileName) Then
            fs.Close()
        End If
    End Sub

    Public Function ExportToTextFile(ByVal FName As String, ByVal OutPutFile As String, _
                            ByVal Sep As String, SelectionOnly As Boolean, Optional ColCount As Long = 1) As Boolean

        Dim WholeLine As String
        Dim RowNdx As Long, ColNdx As Long, wsNameLen As Integer
        Dim StartRow As Long, EndRow As Long
        Dim StartCol As Long, EndCol As Long
        Dim CellValue
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(FName)
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = wb.ActiveSheet

        ' Dim rng As Excel.Range

        Dim PrintAll As Boolean
        Try
            'create text output file
            If (Not File.Exists(OutPutFile)) Then
                File.Create(OutPutFile).Dispose()
            End If
            'sw = New System.IO.StreamWriter(OutPutFile)

            PrintAll = False

            For Each ws In wb.Worksheets
                If SelectionOnly = True Then
                    With ws
                        StartRow = .Cells(1).row
                        StartCol = .Cells(1).Column
                        EndRow = .Cells(.Cells.Count).row
                        EndCol = .Cells(.Cells.Count).Column
                    End With
                Else
                    ws.Activate()

                    With xlApp.ActiveSheet.UsedRange
                        StartRow = .Cells(1).row
                        StartCol = .Cells(1).Column
                        ws.UsedRange.Select()

                        EndRow = ws.UsedRange.Find(What:="*", _
                            SearchDirection:=Excel.XlSearchDirection.xlPrevious, _
                            SearchOrder:=Excel.XlSearchOrder.xlByRows).Row
                        ' If ColCount = 0 Then
                        EndCol = ws.UsedRange.Find(What:="*", _
                            SearchDirection:=Excel.XlSearchDirection.xlPrevious, _
                            SearchOrder:=Excel.XlSearchOrder.xlByColumns).Column
                        ' Else
                        'EndCol = ColCount
                        ' End If
                    End With
                End If
                wsNameLen = 0
                For RowNdx = StartRow To EndRow
                    WholeLine = ""

                    For ColNdx = StartCol To EndCol
                        CellValue = Trim(xlApp.Cells(RowNdx, ColNdx).Value)
                        If CellValue.ToString() = "" Then
                            CellValue = Chr(34) & Chr(34)
                        Else
                            ' CellValue = Trim(xlApp.Cells(RowNdx, ColNdx).Value.ToString())
                        End If
                        WholeLine = WholeLine & CellValue.ToString() & Sep
                    Next ColNdx
                    WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))

                    If Not PrintAll Then
                        If Mid(WholeLine, wsNameLen + 1, 5) <> "Acc #" Then
                            If Mid(WholeLine, wsNameLen + 1, 5) <> Chr(34) & Chr(34) & Sep & Chr(34) & Chr(34) Then
                                'Write to output file
                                If File.Exists(OutPutFile) Then
                                    Using sWrite As StreamWriter = New StreamWriter(OutPutFile, True)
                                        Try
                                            sWrite.WriteLine(WholeLine)
                                            sWrite.Flush()
                                            sWrite.Close()
                                        Catch ex As Exception
                                            MsgBox(ex.Message)
                                        End Try
                                    End Using
                                End If
                            End If
                        End If
                    Else
                        'Write to output file
                        If File.Exists(OutPutFile) Then
                            Using sWrite As StreamWriter = New StreamWriter(OutPutFile, True)
                                Try
                                    sWrite.WriteLine(WholeLine)
                                    sWrite.Flush()
                                    sWrite.Close()
                                Catch ex As Exception
                                    MsgBox(ex.Message)
                                End Try
                            End Using
                        End If
                    End If

                Next RowNdx
            Next ws
            xlApp.DisplayAlerts = False
        Catch ex As Exception
            If ErrMsg <> "" Then
                ErrMsg = ErrMsg & "; " & ex.Message
            Else
                ErrMsg = ex.Message
            End If
            ExportToTextFile = False
            wb.Close(False, FName)
            xlApp.Quit()
            ws = Nothing
            wb = Nothing
            xlApp = Nothing
            Exit Function
        End Try

        wb.Close(False, FName)
        ExportToTextFile = True

        xlApp.Quit()
        ws = Nothing
        wb = Nothing
        xlApp = Nothing

    End Function

    

    Public Sub WriteCSVIntrascriptFile(ByVal outFile As string, ByVal FieldsList As String)
        Dim OutputString As String = ""

        OutputString = OutputString + GetDateNumber(FieldsList, "ChartNo") + "|"
        OutputString = OutputString + GetStr(FieldsList, "AcctNo") + "|"
        OutputString = OutputString + GetDateNumber(FieldsList, "PtDOB") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PtFullName") + "|"
        OutputString = OutputString + GetStr(FieldsList, "Sex") + "|"
        OutputString = OutputString + GetStr(FieldsList, "SSN") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PtAddress") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PtPhone") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PtRace") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PtReligion") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PtMaritalStatus") + "|"
        OutputString = OutputString + GetStr(FieldsList, "FinancialClass") + "|"
        OutputString = OutputString + GetDateNumber(FieldsList, "AdmissionDate") + "|"
        OutputString = OutputString + GetStr(FieldsList, "DischargeDate") + "|"
        OutputString = OutputString + GetStr(FieldsList, "ServiceDate") + "|"
        OutputString = OutputString + GetStr(FieldsList, "RdFullName") + "|"
        OutputString = OutputString + GetStr(FieldsList, "OrderNumber") + "|"
        OutputString = OutputString + GetStr(FieldsList, "VisitNumber") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PatientType") + "|"
        OutputString = OutputString + GetStr(FieldsList, "PatientLocation")

        'Write to output file
        If File.Exists(outFile) Then
            Using sw As StreamWriter = New StreamWriter(outFile, True)
                Try
                    sw.WriteLine(OutputString)
                    sw.Flush()
                    sw.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If


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
    Private Function GetStr(FieldsList As String, KeyName As String)
        Dim Pos As Integer, pos2 As Integer

        Pos = InStr(FieldsList, KeyName)
        If Pos <> 0 Then
            Pos = Pos + Len(KeyName) + 1
            pos2 = InStr(Pos, FieldsList, "|")
            GetStr = "" + Mid(FieldsList, Pos, pos2 - Pos) + ""
        Else
            GetStr = ""
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

End Module
