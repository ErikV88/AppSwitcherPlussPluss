Imports System.Runtime.InteropServices

Public Class WindowFinder

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function

    ''' <summary>
    ''' Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.
    ''' </summary>
    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True)>
    Private Shared Function FindWindowByCaption(ZeroOnly As IntPtr, lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInt32, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Const WM_CLOSE As UInt32 = &H10

    Public Shared Function FindWindow(caption As String) As IntPtr
        If caption.Contains("Skype") Then
            Dim procs As Process() = Process.GetProcessesByName("Skype")
            Return procs(0).MainWindowHandle
        End If


        Return FindWindowByCaption(IntPtr.Zero, caption)
    End Function

End Class
