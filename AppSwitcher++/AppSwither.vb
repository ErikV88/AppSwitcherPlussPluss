
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Partial Public Class AppSwither
    Inherits Form

    Public AppName As String
    Public WindowsC As New List(Of WindowControl)
    Dim vCurrnetWindow As Integer = 0
    Private thumb As IntPtr

#Region "Win32 delicate functions"

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Sub SwitchToThisWindow(hWnd As IntPtr, turnOn As Boolean)
    End Sub


    <DllImport("user32.dll")>
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer,
    uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowLongA(hWnd As IntPtr, nIndex As Integer) As ULong
    End Function

    <DllImport("user32.dll")>
    Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsCallback, lParam As Integer) As Integer
    End Function
    Private Delegate Function EnumWindowsCallback(hwnd As IntPtr, lParam As Integer) As Boolean

    <DllImport("user32.dll")>
    Public Shared Sub GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function GetActiveWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(hWnd As Integer, msg As IntPtr, wParam As Integer, lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ShowWindowAsync(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Public Shared Function GetClassLongPtr(hWnd As IntPtr, nIndex As Integer) As IntPtr
        If IntPtr.Size > 4 Then
            Return GetClassLongPtr64(hWnd, nIndex)
        Else
            Return New IntPtr(GetClassLongPtr32(hWnd, nIndex))
        End If
    End Function

    <DllImport("user32.dll", EntryPoint:="GetClassLong")>
    Public Shared Function GetClassLongPtr32(hWnd As IntPtr, nIndex As Integer) As UInteger
    End Function

    <DllImport("user32.dll", EntryPoint:="GetClassLongPtr")>
    Public Shared Function GetClassLongPtr64(hWnd As IntPtr, nIndex As Integer) As IntPtr
    End Function
#End Region

#Region "Constants"
    Public Const GCL_HICONSM As Integer = -34
    Public Const GCL_HICON As Integer = -14
    Public Const ICON_SMALL As Integer = 0
    Public Const ICON_BIG As Integer = 1
    Public Const ICON_SMALL2 As Integer = 2
    Public Const WM_GETICON As Integer = &H7F
    Public Const GWL_STYLE As Integer = -16
    Public Const DWM_TNP_VISIBLE As Integer = &H8
    Public Const DWM_TNP_OPACITY As Integer = &H4
    Public Const DWM_TNP_RECTDESTINATION As Integer = &H1
    Public Const WS_VISIBLE As ULong = &H10000000L
    Public Const WS_BORDER As ULong = &H800000L
    Public Const TARGETWINDOW As ULong = WS_BORDER Or WS_VISIBLE
    Public Const WM_SYSCOMMAND As Integer = &H112
    Public Const SC_CLOSE As Integer = &HF060
    Private Const WM_CLOSE As UInt32 = &H10
    Private Const SW_HIDE As Integer = 0
    Private Const SW_SHOWNORMAL As Integer = 1
    Private Const SW_SHOWMINIMIZED As Integer = 2
    Private Const SW_SHOWMAXIMIZED As Integer = 3
    Private Const SW_SHOWNOACTIVATE As Integer = 4
    Private Const SW_RESTORE As Integer = 9
    Private Const SW_SHOWDEFAULT As Integer = 10
    Shared ReadOnly HWND_TOPMOST As New IntPtr(-1)
    Const SWP_NOSIZE As UInt32 = &H1
    Const SWP_NOMOVE As UInt32 = &H2
    Const SWP_SHOWWINDOW As UInt32 = &H40
#End Region

#Region "WIN32 helper functions"
    Public Shared Sub WIN32Close(handle As IntPtr)
        SendMessage(handle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)
    End Sub

    Public Shared Sub WIN32bringToFront(handle As IntPtr)
        SwitchToThisWindow(handle, True)

        SetWindowPos(handle, HWND_TOPMOST, 0, 0, 0, 0,
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End Sub

    Public Function GetAppIcon(hwnd As IntPtr) As Icon
        Dim iconHandle As IntPtr = SendMessage(hwnd, WM_GETICON, ICON_SMALL2, 0)
        If iconHandle = IntPtr.Zero Then
            iconHandle = SendMessage(hwnd, WM_GETICON, ICON_SMALL, 0)
        End If
        If iconHandle = IntPtr.Zero Then
            iconHandle = SendMessage(hwnd, WM_GETICON, ICON_BIG, 0)
        End If
        If iconHandle = IntPtr.Zero Then
            iconHandle = GetClassLongPtr(hwnd, GCL_HICON)
        End If
        If iconHandle = IntPtr.Zero Then
            iconHandle = GetClassLongPtr(hwnd, GCL_HICONSM)
        End If

        If iconHandle = IntPtr.Zero Then
            Return Nothing
        End If

        Dim icn As Icon = Icon.FromHandle(iconHandle)

        Return icn
    End Function
#End Region

    Public Sub New(ByVal pAppName As String)
        InitializeComponent()
        AppName = pAppName

        Me.BringToFront()
        Me.Activate()
    End Sub

    Public Sub RefreshWindows()
        Dim vNumberOfRows As Integer
        Dim vNumberOfCells As Integer

        If Windows.Count < 4 Then
            vNumberOfRows = 1
            vNumberOfCells = lstWindows.Items.Count
        Else
            vNumberOfRows = Math.Ceiling(lstWindows.Items.Count / 4)
            vNumberOfCells = 4
        End If

        Dim vWindowHeight As Integer = 0
        Dim vWindowWidth As Integer = 0
        Dim vRowTeller As Integer = 1

        For row As Integer = 1 To vNumberOfRows
            For cell As Integer = 1 To vNumberOfCells
                If lstWindows.Items.Count >= vCurrnetWindow Then
                    If vCurrnetWindow + 1 <= lstWindows.Items.Count Then
                        Dim vWindow As New WindowControl With {.Name = String.Format("User{0}Control{1}", vNumberOfRows, vNumberOfCells), .Window = lstWindows.Items(vCurrnetWindow)}
                        vWindow.Left = vWindow.Width * (cell - 1)
                        If (row - 1) = 0 Then
                            vWindowWidth = vWindowWidth + vWindow.Width
                        End If

                        If cell = 1 Then
                            vWindowHeight = vWindowHeight + vWindow.Height
                        End If

                        vWindow.Top = vWindow.Height * (row - 1)
                        Me.Panel2.Controls.Add(vWindow)
                    End If
                    vCurrnetWindow = vCurrnetWindow + 1
                End If
            Next
        Next

        Dim vWidthMargin As Integer = Panel1.Padding.Left + Panel1.Padding.Right
        Dim vHeightMargin As Integer = Panel1.Padding.Top + (Panel1.Padding.Bottom * 3)
        vWindowWidth = vWindowWidth + vWidthMargin
        vWindowHeight = vWindowHeight + vHeightMargin
        Me.Width = vWindowWidth
        Me.Height = vWindowHeight
    End Sub

    Private Sub AppSwither2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Me.CenterToScreen()

        GetWindows(Me)
        RefreshWindows()

        Me.Opacity = 0.8

        Me.BackColor = Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.BringToFront()
        Me.Focus()
        WIN32bringToFront(Me.Handle)

    End Sub


    Private FormIndex As Integer = 0
    Private SelectedWindow As Integer = 0
    Public Sub AppSwither_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Right Then

            For Each c In Panel2.Controls
                c.BorderStyle = BorderStyle.None
            Next

            If FormIndex + 1 <= Panel2.Controls.Count Then
                CType(Panel2.Controls.Item(FormIndex), WindowControl).BorderStyle = BorderStyle.Fixed3D
                FormIndex = FormIndex + 1
            Else
                FormIndex = 0
                If FormIndex + 1 <= Panel2.Controls.Count Then
                    CType(Panel2.Controls.Item(FormIndex), WindowControl).BorderStyle = BorderStyle.Fixed3D
                End If

                FormIndex = 1
            End If

            Return
        End If
    End Sub

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If lstWindows.Items.Count < 1 Then Return
            Dim vHander = WindowFinder.FindWindow(CType(Panel2.Controls(FormIndex - 1), WindowControl).Window.Title)  ' GetWindowByName(DirectCast(lstWindows.Items(vIndex), Window).Title)
            SetForegroundWindow(vHander)
            ShowWindowAsync(vHander, SW_SHOWMAXIMIZED)

            e.Cancel = False
        Catch ex As Exception

        End Try
    End Sub

    Public Sub AppSwither_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp

        If AppSwitcherContext.getSynonymKey(Keys.Modifiers) And Not Config.useXkey Then
            AppSwitcherContext.CloseApp(Me)
        End If


        If e.KeyCode = Keys.Enter And 1 = 0 Then
            Dim vHander = WindowFinder.FindWindow(DirectCast(lstWindows.Items(If(FormIndex = 1, 1, 0)), Window).Title)  ' GetWindowByName(DirectCast(lstWindows.Items(vIndex), Window).Title)
            SetForegroundWindow(vHander)
            ShowWindowAsync(vHander, SW_SHOWMAXIMIZED)
            Me.DialogResult = DialogResult.OK

            If Not Config.useXkey Then Return
            WIN32Close(Me.Handle)
        End If
    End Sub

    Public lstWindows As New ComboBox
    <DllImport("dwmapi.dll")>
    Private Shared Function DwmRegisterThumbnail(dest As IntPtr, src As IntPtr, ByRef thumb As IntPtr) As Integer
    End Function


#Region "Retrieve list of windows"
    Private Windows As List(Of Window)
    Private Sub GetWindows(sender As Form)
        Windows = New List(Of Window)()
        EnumWindows(AddressOf Callback, 0)

        lstWindows.Items.Clear()
        For Each w As Window In Windows
            If w.ToString().Contains(AppName) Then
                Handler = sender.Handle
                lstWindows.Items.Add(w)

            End If
        Next
    End Sub
    Private Handler As IntPtr
    Private Function Callback(hwnd As IntPtr, lParam As Integer) As Boolean
        If Handler <> hwnd AndAlso (GetWindowLongA(hwnd, GWL_STYLE) And TARGETWINDOW) = TARGETWINDOW Then
            Dim sb As New StringBuilder(100)
            GetWindowText(hwnd, sb, sb.Capacity)

            Dim t As New Window()

            t.Handle = hwnd
            t.Title = sb.ToString()
            Windows.Add(t)
        End If

        Return True
        'continue enumeration
    End Function

#End Region
End Class


