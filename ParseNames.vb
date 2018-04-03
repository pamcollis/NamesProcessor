Option Explicit On
Option Compare Text

Module ParseNames

    Public PrefixArray() As String
    Public SuffixArray() As String
    Public ONealMcDonald() As String
    Public Function ParseLFMName(FullName As String, DoCase As String) As Object
        Dim WorkName, FirstName, LastName, MiddleName, WorkPrefSuff, Prefix, Suffix As String
        Dim i As Integer, Pos As Integer, pos2 As Integer

        WorkName = "" : FirstName = "" : LastName = "" : MiddleName = "" : WorkPrefSuff = "" : Prefix = "" : Suffix = ""
        If PrefixArray Is Nothing Then
            InitializeParser()
        End If
        '***************************************************************************
        'Get rid of all extraneous characters
        WorkName = Translate(FullName, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,-/'1234567890~`@#$%^&*()_+={[}]|\:;""<>.?¡¢£¤¥¦§¨©ª­®¯°±²³´µ·¸¹º¼½¾¿", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,- '")
        '***************************************************************************
        'Get rid of all double spaces
        While InStr(WorkName, "  ")
            WorkName = Left(WorkName, InStr(WorkName, "  ")) + Mid(WorkName, InStr(WorkName, "  ") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Get rid of all double commas
        While InStr(WorkName, ",,")
            WorkName = Left(WorkName, InStr(WorkName, ",,")) + Mid(WorkName, InStr(WorkName, ",,") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Pull out all of the prefixes
        Pos = 0
        While InStr(Pos + 1, WorkName, " ") <> 0 Or Pos <= Len(WorkName)
            If InStr(Pos + 1, WorkName, ",") <> 0 And InStr(Pos + 1, WorkName, ",") < InStr(Pos + 1, WorkName, " ") Then
                pos2 = InStr(Pos + 1, WorkName, ",")
                If pos2 = 0 Then
                    pos2 = Len(WorkName) + 1
                End If
                WorkPrefSuff = Mid(WorkName, Pos + 1, pos2 - (Pos + 1))
                If Pos <> 0 Then  'Cant be a prefix in first word
                    For i = 0 To UBound(PrefixArray)
                        If WorkPrefSuff = PrefixArray(i) Then
                            Prefix = Trim(Prefix + " " + PrefixArray(i))
                            WorkName = Left(WorkName, Pos) + Right(WorkName, Len(WorkName) - pos2 + 1)
                            WorkPrefSuff = "~"
                            Exit For
                        End If
                    Next i
                End If
            Else
                pos2 = InStr(Pos + 1, WorkName, " ")
                If pos2 = 0 Then
                    pos2 = Len(WorkName) + 1
                End If
                WorkPrefSuff = Mid(WorkName, Pos + 1, pos2 - (Pos + 1))
                If Pos <> 0 Then  'Cant be a prefix in first word
                    For i = 0 To UBound(PrefixArray)
                        If WorkPrefSuff = PrefixArray(i) Then
                            Prefix = Trim(Prefix + " " + PrefixArray(i))
                            WorkName = Left(WorkName, Pos) + Right(WorkName, Len(WorkName) - pos2 + 1)
                            WorkPrefSuff = "~"
                            Exit For
                        End If
                    Next i
                End If
            End If
            If WorkPrefSuff <> "~" Then
                Pos = pos2
            End If
        End While
        '***************************************************************************
        'Get rid of all double spaces
        While InStr(WorkName, "  ")
            WorkName = Left(WorkName, InStr(WorkName, "  ")) + Mid(WorkName, InStr(WorkName, "  ") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Get rid of all double commas
        While InStr(WorkName, ",,")
            WorkName = Left(WorkName, InStr(WorkName, ",,")) + Mid(WorkName, InStr(WorkName, ",,") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Pull out all of the suffixes
        Pos = 0
        While InStr(Pos + 1, WorkName, " ") <> 0 Or Pos <= Len(WorkName)
            If InStr(Pos + 1, WorkName, ",") <> 0 And InStr(Pos + 1, WorkName, ",") < InStr(Pos + 1, WorkName, " ") Then
                pos2 = InStr(Pos + 1, WorkName, ",")
                If pos2 = 0 Then
                    pos2 = Len(WorkName) + 1
                End If
                WorkPrefSuff = Mid(WorkName, Pos + 1, pos2 - (Pos + 1))
                If Pos <> 0 Then  'Cant be a suffix in first word
                    For i = 0 To UBound(SuffixArray)
                        If WorkPrefSuff = SuffixArray(i) Then
                            Suffix = Trim(Suffix + " " + SuffixArray(i))
                            WorkName = Left(WorkName, Pos) + Right(WorkName, Len(WorkName) - pos2 + 1)
                            WorkPrefSuff = "~"
                            Exit For
                        End If
                    Next i
                End If
            Else
                pos2 = InStr(Pos + 1, WorkName, " ")
                If pos2 = 0 Then
                    pos2 = Len(WorkName) + 1
                End If
                WorkPrefSuff = Mid(WorkName, Pos + 1, pos2 - (Pos + 1))
                If Pos <> 0 Then  'Cant be a suffix in first word
                    For i = 0 To UBound(SuffixArray)
                        If WorkPrefSuff = SuffixArray(i) Then
                            Suffix = Trim(Suffix + " " + SuffixArray(i))
                            WorkName = Left(WorkName, Pos) + Right(WorkName, Len(WorkName) - pos2 + 1)
                            WorkPrefSuff = "~"
                            Exit For
                        End If
                    Next i
                End If
            End If
            If WorkPrefSuff <> "~" Then
                Pos = pos2
            End If
        End While
        '***************************************************************************
        'Get rid of all double spaces
        While InStr(WorkName, "  ")
            WorkName = Left(WorkName, InStr(WorkName, "  ")) + Mid(WorkName, InStr(WorkName, "  ") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Get rid of all extra commas
        If InStr(InStr(WorkName, ",") + 1, WorkName, ",") <> 0 Then
            WorkName = Left(WorkName, InStr(WorkName, ",")) + Translate(Mid(WorkName, InStr(WorkName, ",") + 1), "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
        End If
        '***************************************************************************
        'Pull out name elements
        Pos = InStr(WorkName, ",")
        If Pos = 0 Then
            Pos = InStr(WorkName, " ")
            If Pos = 0 Then
                LastName = Trim(WorkName)
            Else
                LastName = Trim(Left(WorkName, Pos - 1))
                pos2 = InStr(Pos + 1, WorkName, " ")
                While pos2 - Pos = 1
                    Pos = pos2
                    pos2 = InStr(Pos + 1, WorkName, " ")
                End While
                If pos2 = 0 Then
                    pos2 = Len(WorkName) + 1
                    FirstName = Trim(Mid(WorkName, Pos + 1, pos2 - Pos - 1))
                    MiddleName = ""
                Else
                    FirstName = Trim(Mid(WorkName, Pos + 1, pos2 - Pos - 1))
                    MiddleName = Trim(Right(WorkName, Len(WorkName) - pos2))
                End If
            End If
        Else
            LastName = Trim(Left(WorkName, Pos - 1))
            pos2 = InStr(Pos + 1, WorkName, " ")
            While pos2 - Pos = 1
                Pos = pos2
                pos2 = InStr(Pos + 1, WorkName, " ")
            End While
            If pos2 = 0 Then
                pos2 = Len(WorkName) + 1
                FirstName = Trim(Mid(WorkName, Pos + 1, pos2 - Pos - 1))
                MiddleName = ""
            Else
                FirstName = Trim(Mid(WorkName, Pos + 1, pos2 - Pos - 1))
                MiddleName = Trim(Right(WorkName, Len(WorkName) - pos2))
            End If
        End If
        '***************************************************************************
        'Do casing based on parameter
        If DoCase = "P" Then
            FirstName = StrConv(FirstName, vbProperCase)
            MiddleName = StrConv(MiddleName, vbProperCase)
            LastName = StrConv(LastName, vbProperCase)
            For i = 0 To UBound(ONealMcDonald)
                If Left(LastName, Len(ONealMcDonald(i))) = ONealMcDonald(i) Then
                    LastName = Left(LastName, Len(ONealMcDonald(i))) + UCase(Mid(LastName, Len(ONealMcDonald(i)) + 1, 1)) + Mid(LastName, Len(ONealMcDonald(i)) + 2)
                    Exit For
                End If
            Next i
        ElseIf DoCase = "U" Then
            Prefix = StrConv(Prefix, vbUpperCase)
            FirstName = StrConv(FirstName, vbUpperCase)
            MiddleName = StrConv(MiddleName, vbUpperCase)
            LastName = StrConv(LastName, vbUpperCase)
            Suffix = StrConv(Suffix, vbUpperCase)
        ElseIf DoCase = "L" Then
            Prefix = StrConv(Prefix, vbLowerCase)
            FirstName = StrConv(FirstName, vbLowerCase)
            MiddleName = StrConv(MiddleName, vbLowerCase)
            LastName = StrConv(LastName, vbLowerCase)
            Suffix = StrConv(Suffix, vbLowerCase)
        End If
        ParseLFMName = {Prefix, FirstName, MiddleName, LastName, Suffix}

    End Function
    Function ParseFMLName(FullName As String, DoCase As String) As Object

        Dim WorkName, FirstName, LastName, MiddleName, WorkPrefSuff, Prefix, Suffix As String
        Dim i As Integer, n As Integer, Pos As Integer, pos2 As Integer
        WorkName = "" : FirstName = "" : LastName = "" : MiddleName = "" : WorkPrefSuff = "" : Prefix = "" : Suffix = ""
        If PrefixArray Is Nothing Then
            InitializeParser()
        End If
        '***************************************************************************
        'Get rid of all extraneous characters
        WorkName = Translate(FullName, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,-/',1234567890~`@#$%^&*()_+={[}]|\:;""<>.?¡¢£¤¥¦§¨©ª­®¯°±²³´µ·¸¹º¼½¾¿", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz   '")
        '***************************************************************************
        'Get rid of all double spaces
        While InStr(WorkName, "  ")
            WorkName = Left(WorkName, InStr(WorkName, "  ")) + Mid(WorkName, InStr(WorkName, "  ") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Pull out all of the prefixes
        Do While InStr(WorkName, " ") <> 0
            If InStr(WorkName, ",") <> 0 And InStr(WorkName, ",") < InStr(WorkName, " ") Then
                WorkPrefSuff = Mid(WorkName, 1, InStr(WorkName, ",") - 1)
                For i = 0 To UBound(PrefixArray)
                    If WorkPrefSuff = PrefixArray(i) Then
                        Prefix = Trim(Prefix + " " + PrefixArray(i))
                        WorkName = Right(WorkName, Len(WorkName) - InStr(WorkName, ","))
                        WorkPrefSuff = "~"
                        Exit For
                    End If
                Next i
            Else
                WorkPrefSuff = Mid(WorkName, 1, InStr(WorkName, " ") - 1)
                For i = 0 To UBound(PrefixArray)
                    If WorkPrefSuff = PrefixArray(i) Then
                        Prefix = Trim(Prefix + " " + PrefixArray(i))
                        WorkName = Right(WorkName, Len(WorkName) - InStr(WorkName, " "))
                        WorkPrefSuff = "~"
                        Exit For
                    End If
                Next i
            End If
            If WorkPrefSuff <> "~" Then
                Exit Do
            End If
        Loop
        '***************************************************************************
        'Get rid of all double spaces
        While InStr(WorkName, "  ")
            WorkName = Left(WorkName, InStr(WorkName, "  ")) + Mid(WorkName, InStr(WorkName, "  ") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Pull out all of the suffixes
        For n = Len(WorkName) To 1 Step -1
            If Mid(WorkName, n, 1) = " " Then
                WorkPrefSuff = Mid(WorkName, n + 1)
                For i = 0 To UBound(SuffixArray)
                    If WorkPrefSuff = SuffixArray(i) Then
                        Suffix = Trim(SuffixArray(i) + " " + Suffix)
                        WorkName = Left(WorkName, n - 1)
                        WorkPrefSuff = "~"
                        Exit For
                    End If
                Next i
                If WorkPrefSuff <> "~" Then
                    Exit For
                End If
            End If
        Next n
        '***************************************************************************
        'Get rid of all double spaces
        While InStr(WorkName, "  ")
            WorkName = Left(WorkName, InStr(WorkName, "  ")) + Mid(WorkName, InStr(WorkName, "  ") + 2)
        End While
        WorkName = Trim(WorkName)
        '***************************************************************************
        'Pull out name elements
        Pos = InStr(WorkName, " ")
        If Pos = 0 Then
            FirstName = Trim(WorkName)
        Else
            FirstName = Trim(Left(WorkName, Pos - 1))
            pos2 = InStr(Pos + 1, WorkName, " ")
            If pos2 = 0 Then
                MiddleName = ""
                LastName = Trim(Mid(WorkName, Pos + 1))
            Else
                MiddleName = Trim(Mid(WorkName, Pos + 1, pos2 - Pos - 1))
                LastName = Trim(Right(WorkName, Len(WorkName) - pos2))
            End If
        End If
        '***************************************************************************
        'Do casing based on parameter
        If DoCase = "P" Then
            FirstName = StrConv(FirstName, vbProperCase)
            MiddleName = StrConv(MiddleName, vbProperCase)
            LastName = StrConv(LastName, vbProperCase)
            For i = 0 To UBound(ONealMcDonald)
                If Left(LastName, Len(ONealMcDonald(i))) = ONealMcDonald(i) Then
                    LastName = Left(LastName, Len(ONealMcDonald(i))) + UCase(Mid(LastName, Len(ONealMcDonald(i)) + 1, 1)) + Mid(LastName, Len(ONealMcDonald(i)) + 2)
                    Exit For
                End If
            Next i
        ElseIf DoCase = "U" Then
            Prefix = StrConv(Prefix, vbUpperCase)
            FirstName = StrConv(FirstName, vbUpperCase)
            MiddleName = StrConv(MiddleName, vbUpperCase)
            LastName = StrConv(LastName, vbUpperCase)
            Suffix = StrConv(Suffix, vbUpperCase)
        ElseIf DoCase = "L" Then
            Prefix = StrConv(Prefix, vbLowerCase)
            FirstName = StrConv(FirstName, vbLowerCase)
            MiddleName = StrConv(MiddleName, vbLowerCase)
            LastName = StrConv(LastName, vbLowerCase)
            Suffix = StrConv(Suffix, vbLowerCase)
        End If
        ParseFMLName = {Prefix, FirstName, MiddleName, LastName, Suffix}

    End Function

    Function CheckSuffix(WorkName As String) As Boolean
        Dim i As Integer

        CheckSuffix = False
        For i = 0 To UBound(SuffixArray)
            If WorkName = SuffixArray(i) Then
                CheckSuffix = True
            End If
        Next i

    End Function
    Public Sub InitializeParser()
        'Initializes common variables for the parser
        PrefixArray = {"and", "Bro", "Brother", "Capt", "Captain", "Chm", "Co", "Col", _
                            "Colonel", "Corporal", "Cpl", "Cpt", "Den", "Dep", "Deputy", _
                            "Doctor", "Dr", "Father", "Fr", "Gen", "General", _
                            "Hon", "Honorable", "Leutenant", "Lt", "Ltl", "Mme", "Madam", "Maj", _
                            "Major", "Mdm", "Miss", "Mister", "Mr", "Mrs", "Ms", _
                            "Pastor", "Prf", "Pfc", "Pivate", "Prof", "Preoessor", _
                            "Prs", "Pvt", "Rab", "Rabbi", "Rep", "Representative", "Rev", _
                            "Reverend", "Rtr", "Sen", "Senator", "SfcC", "Sgt", _
                            "Sherrif", "Sir", "Sister", "Ssg", "Ssgt"}

        SuffixArray = {"ADJ", "ACP", "AP", "ARNP", "ARNPC", "ASOTP", "ATRBC", _
                            "BA", "BCD", "BS", "BSC", "BSW", "CAC", "CADC", "CARN", _
                            "CCJS", "CCS", "CEAP", "CEO", "CHAIRMAN", "CHB", "CHM", _
                            "CMA", "CMD", "CNM", "CNS", "CPA", "CRC", "CRNA", _
                            "CSMS", "CST", "CSW", "CVRN", "DC", "DDS", "DM", _
                            "DMD", "DMIN", "DO", "DP", "DPM", "DSW", "DTR", "DVM", _
                            "ED", "EdD", "EDM", "EDS", "Esq", "EVP", _
                            "FACC", "FACG", "FACP", "FACS", "Family", "Fellow", _
                            "FNP", "II", "III", "Intern", "IV", "JD", "Jr", _
                            "LAC-PH", "LCCA", "LCDC", "LCSW", "LMFT", "LMHC", "LMSW", _
                            "LMT", "LPC", "LPN", "LSW", "LSWA", "MA", "MAC", "MB", _
                            "MCM", "MD", "M.D.", "MED", "MFT", "MHS", "MS", "MSN", "MSW", _
                            "NCAC", "NCC", "OD", "OMD", "ORT", "PA", "PAC", _
                            "PhD", "Psy", "Psychiatrist", "Psychologist", "Psychology", _
                            "PsyD", "PT", "Resident", "RN", "RNA", "RNC", "RPT", _
                            "RSOTP", "Sr", "SVP", "VNC", "VP"}
        ONealMcDonald = {"O'", "Mc", "D'", "L'"}
    End Sub
    Public Function Translate(InString As String, Chars1 As String, Chars2 As String) As String
        'Returns Instring with all occurrences of characters in Chars1 replaced by the corresponding
        'character (based on position) of Chars2.  If the character in Chars1 does not have a
        'corresponding character in Chars2 it is removed.

        Dim TempString As String
        Dim i As Integer
        Dim n As Integer
        Dim NewChar As String

        If InString = "" Then
            Translate = InString
        Else
            TempString = ""
            For i = 1 To Len(InString)
                NewChar = Mid(InString, i, 1)  'If not in first char set leave as is
                For n = 1 To Len(Chars1)
                    If NewChar = Mid(Chars1, n, 1) Then  'Found in first char set
                        NewChar = Mid(Chars2, n, 1)  'Replace from second set or blank if not there
                    End If
                Next n
                TempString = TempString & NewChar
            Next i
            Translate = TempString
        End If

    End Function
End Module
