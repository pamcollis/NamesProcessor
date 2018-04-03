Imports System.Runtime.InteropServices
Imports System.Text


Public Class IniFile


    <DllImport("kernel32")> _
    Private Shared Function GetPrivateProfileString(section As String, key As String, def As String, retVal As StringBuilder, size As Integer, filePath As String) As Integer
    End Function
    Public Function IniReadValue(Section As String, Key As String, iniPath As String) As String
        Dim temp As New StringBuilder(255)
        Dim i As Integer = GetPrivateProfileString(Section, Key, "", temp, 255, iniPath)
        Return temp.ToString()

    End Function
End Class

