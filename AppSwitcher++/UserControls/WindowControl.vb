Imports System.Runtime.InteropServices
Imports System.Text

Public Class WindowControl
    Public Property AppIcon As Image
        Set(value As Image)
            PictureBox1.Image = value
        End Set
        Get
            Return PictureBox1.Image
        End Get
    End Property

    Public Property Window As Window


#Region "Propertys"
    Private Property FormIndex As Integer = 1
    Public Property thumb As IntPtr
    Public TrackBar1 As New TrackBar
    Public lstWindows As New ComboBox
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
    Private Sub UpdateThumb()

        If thumb <> IntPtr.Zero Then

            Dim size As PSIZE
            DwmQueryThumbnailSourceSize(thumb, size)
            Dim vMargin As Integer = 20
            Dim props As New DWM_THUMBNAIL_PROPERTIES()
            props.fVisible = True
            props.dwFlags = DWM_TNP_VISIBLE Or DWM_TNP_RECTDESTINATION Or DWM_TNP_OPACITY
            props.opacity = CByte(TrackBar1.Value)
            props.rcDestination = New Rect(Me.Left + vMargin, vMargin + Me.Top + SplitContainer1.Panel1.Height, Me.Right + vMargin, Me.Bottom + vMargin)

            If size.x < SplitContainer1.Width Then
                props.rcDestination.Right = props.rcDestination.Left + size.x
            End If


            If size.y < SplitContainer1.Panel2.Height Then
                props.rcDestination.Bottom = props.rcDestination.Top + size.y
            End If


            DwmUpdateThumbnailProperties(thumb, props)
        End If
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
    Private Sub UserControl1_Load(sender As UserControl, e As EventArgs) Handles MyBase.Load
        If thumb <> IntPtr.Zero Then
            DwmUnregisterThumbnail(thumb)
        End If

        Dim vRand As Integer = New Random(1000).Next()
        SplitContainer1.Name = String.Format("SplitContainer1{0}", vRand)

        TrackBar1.Value = TrackBar1.Maximum

        RegisterWindow()
        UpdateThumb()
        If Me.Window IsNot Nothing Then
            AppTitle1.Text = Me.Window.Title
            If GetAppIcon(Me.Window.Handle) IsNot Nothing Then
                PictureBox1.Image = GetAppIcon(Me.Window.Handle).ToBitmap
            End If
        End If


    End Sub

    Public Sub RegisterWindow()
        EnumWindows(AddressOf Callback, 0)
        Dim i As Integer = DwmRegisterThumbnail(Me.ParentForm.Handle, Window.Handle, thumb)
    End Sub
#Region "Retrieve list of windows"

    Private Function Callback(hwnd As IntPtr, lParam As Integer) As Boolean
        Return True
        If Me.Handle <> hwnd AndAlso (GetWindowLongA(hwnd, GWL_STYLE) And TARGETWINDOW) = TARGETWINDOW Then
            Dim sb As New StringBuilder(100)
            GetWindowText(hwnd, sb, sb.Capacity)

            Dim t As New Window()
            Me.Window.Handle = hwnd
            Me.Window.Title = sb.ToString()
            ' Windows.Add(t)
        End If

        Return True
        'continue enumeration
    End Function

    Private Sub UserControl1_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        CType(Me.ParentForm, AppSwither).AppSwither_KeyUp(sender, e)
    End Sub

#End Region
End Class
