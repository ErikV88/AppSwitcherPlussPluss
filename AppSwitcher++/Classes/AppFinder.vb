Imports Microsoft.Win32

Public Module AppFinder
    Public Function Chrome() As String
        Dim vValue As String = ""
        Try
            vValue = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe", True).GetValue("").ToString
        Catch ex As Exception
        End Try

        Return vValue
    End Function


    Public Function Firefox() As String
        Dim vValue As String = ""
        Try
            vValue = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\firefox.exe", True).GetValue("").ToString
        Catch ex As Exception
        End Try

        Return vValue
    End Function

End Module
