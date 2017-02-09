
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
#Region "Propertys"
    Private Property FormIndex As Integer = 1
    Private thumb As IntPtr
    Public Property AppName As String
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

#Region "DWM functions"

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmRegisterThumbnail(dest As IntPtr, src As IntPtr, ByRef thumb As IntPtr) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmUnregisterThumbnail(thumb As IntPtr) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmQueryThumbnailSourceSize(thumb As IntPtr, ByRef size As PSIZE) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmUpdateThumbnailProperties(hThumb As IntPtr, ByRef props As DWM_THUMBNAIL_PROPERTIES) As Integer
    End Function

#End Region

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


#Region "Helper Function"

#End Region

    Public Sub New(ByVal pAppName As String)
        InitializeComponent()
        AppName = pAppName

        Me.BringToFront()
        Me.Activate()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs)
        GetWindows()
    End Sub

    Public Sub CloseForm()
        Me.Dispose(True)
    End Sub

#Region "Retrieve list of windows"
    Private windows As List(Of Window)

    Private Sub GetWindows()
        windows = New List(Of Window)()
        EnumWindows(AddressOf Callback, 0)

        lstWindows.Items.Clear()
        For Each w As Window In windows
            If w.ToString().Contains(AppName) Then
                Dim i As Integer = DwmRegisterThumbnail(Me.Handle, w.Handle, thumb)

                lstWindows.Items.Add(w)
            End If
        Next
    End Sub

    Private Function Callback(hwnd As IntPtr, lParam As Integer) As Boolean
        If Me.Handle <> hwnd AndAlso (GetWindowLongA(hwnd, GWL_STYLE) And TARGETWINDOW) = TARGETWINDOW Then
            Dim sb As New StringBuilder(100)
            GetWindowText(hwnd, sb, sb.Capacity)

            Dim t As New Window()
            t.Handle = hwnd
            t.Title = sb.ToString()
            windows.Add(t)
        End If

        Return True
        'continue enumeration
    End Function

#End Region

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        If thumb <> IntPtr.Zero Then
            DwmUnregisterThumbnail(thumb)
        End If

        GetWindows()
    End Sub

    Private Sub lstWindows_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstWindows.SelectedIndexChanged
        Return
        Dim w As Window = DirectCast(lstWindows.SelectedItem, Window)

        If thumb <> IntPtr.Zero Then
            DwmUnregisterThumbnail(thumb)
        End If

        Dim i As Integer = DwmRegisterThumbnail(Me.Handle, w.Handle, thumb)

        If i = 0 Then
            UpdateThumb()
        End If
    End Sub

#Region "Update thumbnail properties"
    Private Sub UpdateThumb2()
        If thumb <> IntPtr.Zero Then
            Dim size As PSIZE
            DwmQueryThumbnailSourceSize(thumb, size)

            Dim props As New DWM_THUMBNAIL_PROPERTIES()

            props.fVisible = True
            props.dwFlags = DWM_TNP_VISIBLE Or DWM_TNP_RECTDESTINATION Or DWM_TNP_OPACITY
            props.opacity = CByte(TrackBar1.Value)
            props.rcDestination = New Rect(SplitContainer1.Left, SplitContainer1.Top + SplitContainer1.Panel1.Height, SplitContainer1.Right, SplitContainer1.Bottom)

            If size.x < SplitContainer1.Width Then
                props.rcDestination.Right = props.rcDestination.Left + size.x
            End If

            If size.y < SplitContainer1.Panel2.Height Then
                props.rcDestination.Bottom = props.rcDestination.Top + size.y
            End If

            DwmUpdateThumbnailProperties(thumb, props)
        End If
    End Sub

    Private Sub UpdateThumb()
        If thumb <> IntPtr.Zero Then
            Dim size As PSIZE
            DwmQueryThumbnailSourceSize(thumb, size)

            Dim props As New DWM_THUMBNAIL_PROPERTIES()

            props.fVisible = True
            props.dwFlags = DWM_TNP_VISIBLE Or DWM_TNP_RECTDESTINATION Or DWM_TNP_OPACITY
            props.opacity = CByte(TrackBar1.Value)
            props.rcDestination = New Rect(SplitContainer2.Left, SplitContainer2.Top + SplitContainer2.Panel1.Height, SplitContainer2.Right, SplitContainer2.Bottom)

            If size.x < SplitContainer2.Width Then
                props.rcDestination.Right = props.rcDestination.Left + size.x
            End If

            If size.y < SplitContainer2.Panel2.Height Then
                props.rcDestination.Bottom = props.rcDestination.Top + size.y
            End If

            DwmUpdateThumbnailProperties(thumb, props)
        End If
    End Sub

    Private Sub opacity_Scroll(sender As Object, e As EventArgs)
        UpdateThumb()
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs)
        UpdateThumb()
    End Sub

    Private Sub AppSwither_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Me.CenterToScreen()

        If thumb <> IntPtr.Zero Then
            DwmUnregisterThumbnail(thumb)
        End If

        GetWindows()

        Me.Opacity = 0.8

        Me.BackColor = Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.BringToFront()
        Me.Focus()

        UpdateThumb()
        SplitContainer2.BorderStyle = BorderStyle.None
        SplitContainer1.BorderStyle = BorderStyle.Fixed3D

        TrackBar1.Value = TrackBar1.Maximum

        Dim vListOfProgram As List(Of Program) = RegestryContext.Apps

        If lstWindows.Items.Count = 0 Then
            Dim vItem = (From p In vListOfProgram
                         Where p.Name = AppName).FirstOrDefault

            If vItem IsNot Nothing Then
                Try
                    Process.Start(vItem.Path)
                Catch ex As Exception

                End Try
            End If

            Return
        End If

        If lstWindows.Items.Count > 0 Then
            Dim w As Window = DirectCast(lstWindows.Items(0), Window)
            AppTitle1.Text = w.Title
            PictureBox1.Image = GetAppIcon(w.Handle).ToBitmap

            Dim i As Integer = DwmRegisterThumbnail(Me.Handle, w.Handle, thumb)
            UpdateThumb2()
        End If

        If lstWindows.Items.Count > 1 Then
            Dim w2 As Window = DirectCast(lstWindows.Items(1), Window)
            AppTitle2.Text = w2.Title
            If GetAppIcon(w2.Handle) IsNot Nothing Then
                PictureBox4.Image = GetAppIcon(w2.Handle).ToBitmap
            End If

            Dim i2 As Integer = DwmRegisterThumbnail(Me.Handle, w2.Handle, thumb)

            UpdateThumb()
        Else
            Me.Width = Me.Width / 2
        End If
        WIN32bringToFront(Me.Handle)
    End Sub

    Public Sub AppSwither_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Right Then
            If FormIndex = 0 Then
                SplitContainer2.BorderStyle = BorderStyle.None
                SplitContainer1.BorderStyle = BorderStyle.Fixed3D

                FormIndex = 1
            Else
                SplitContainer2.BorderStyle = BorderStyle.Fixed3D
                SplitContainer1.BorderStyle = BorderStyle.None
                FormIndex = 0
            End If
            Return
        End If
    End Sub

    Private Sub AppSwither_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        If e.Control AndAlso e.KeyCode = Keys.A Then
            AppSwitcherContext.CloseApp(Me)
        End If

        If e.KeyCode = Keys.Enter And 1 = 0 Then
            Dim vHander = WindowFinder.FindWindow(DirectCast(lstWindows.Items(If(FormIndex = 1, 1, 0)), Window).Title)  ' GetWindowByName(DirectCast(lstWindows.Items(vIndex), Window).Title)
            SetForegroundWindow(vHander)
            ShowWindowAsync(vHander, SW_SHOWMAXIMIZED)
            Me.DialogResult = DialogResult.OK

            WIN32Close(Me.Handle)
        End If
    End Sub

    Private Function GetWindowByName(name As String) As IntPtr
        Dim hWnd As IntPtr = IntPtr.Zero

        Dim procs As Process() = Process.GetProcessesByName("chrome")
        For Each proc As System.Diagnostics.Process In procs
            If proc.ProcessName.Contains("chrome") Then
                Dim a As String = ""
            End If
        Next

        For Each pList As Process In Process.GetProcesses()

            If pList.MainWindowTitle.ToString.Contains(name) Then
                Return pList.MainWindowHandle
            End If
        Next
        Return hWnd
    End Function

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If lstWindows.Items.Count < 1 Then Return
            Dim vHander = WindowFinder.FindWindow(DirectCast(lstWindows.Items(If(FormIndex = 1, 0, 1)), Window).Title)  ' GetWindowByName(DirectCast(lstWindows.Items(vIndex), Window).Title)
            SetForegroundWindow(vHander)
            ShowWindowAsync(vHander, SW_SHOWMAXIMIZED)

            e.Cancel = False
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        FormIndex = 0
        SplitContainer2.BorderStyle = BorderStyle.None
        SplitContainer1.BorderStyle = BorderStyle.Fixed3D
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        FormIndex = 1
        SplitContainer2.BorderStyle = BorderStyle.Fixed3D
        SplitContainer1.BorderStyle = BorderStyle.None
    End Sub

#End Region
End Class

Friend Class Window
        Public Title As String
        Public Handle As IntPtr

        Public Overrides Function ToString() As String
            Return Title
        End Function
    End Class

#Region "Interop structs"

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure DWM_THUMBNAIL_PROPERTIES
        Public dwFlags As Integer
        Public rcDestination As Rect
        Public rcSource As Rect
        Public opacity As Byte
        Public fVisible As Boolean
        Public fSourceClientAreaOnly As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure Rect
        Friend Sub New(left__1 As Integer, top__2 As Integer, right__3 As Integer, bottom__4 As Integer)
            Left = left__1
            Top = top__2
            Right = right__3
            Bottom = bottom__4
        End Sub

        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure PSIZE
        Public x As Integer
        Public y As Integer
    End Structure

#End Region
