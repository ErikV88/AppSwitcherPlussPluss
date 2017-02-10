Imports System.Text
Imports System.Runtime.InteropServices
Imports OutLook = Microsoft.Office.Interop.Outlook
Imports Office = Microsoft.Office.Core
Imports Interop
Imports Interop.SKYPE4COMLib
Imports PIEHid32Net
Imports Limilabs.Client.IMAP
Imports Limilabs.Mail
Imports System.IO

Public Class Main
    Implements PIEHid32Net.PIEDataHandler
    Implements PIEHid32Net.PIEErrorHandler

    <DllImport("user32.dll")>
    Public Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As Integer, vlc As Integer) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

#Region "Propertys and Delegates"
    Dim devices() As PIEHid32Net.PIEDevice
    Private Property selecteddevice As Integer
    Dim cbotodevice(127) As Integer 'max # of devices = 128 
    Dim wdata() As Byte = New Byte() {} 'write data buffer
    Private Property saveabsolutetime As Long
    Private Property c As Control
    Private Property mouseon As Boolean = False
    Private Property EnumerationSuccess As Boolean
    Private Property FormKeys As New BtnKeys
    Private Property Form As AppSwither
    Private Property FormsPool As New List(Of Form)
    Private Property CboDevices As New List(Of String)

    Delegate Sub SetTextCallback(ByVal [text] As String)
#End Region

#Region "Helper Functions"
    Public Sub CloseForms()
        For Each f In FormsPool
            f.Close()
            f.Dispose()
        Next
    End Sub
#End Region

#Region "X-Keys Events"

    Public Sub HandlePIEHidData(ByVal data() As Byte, ByVal sourceDevice As PIEHid32Net.PIEDevice, ByVal perror As Integer) Implements PIEHid32Net.PIEDataHandler.HandlePIEHidData
        If Not Config.useXkey Then
            Return
        End If

        'check the sourceDevice and make sure it is the same device as selected in CboDevice   
        If sourceDevice.Pid = devices(selecteddevice).Pid Then
            Me.Invoke(DirectCast(Sub()
                                     Me.BringToFront()
                                     Me.Focus()

                                 End Sub, Action))

            FormKeys.VS.Normal = CByte(data(3) And 8)
            If FormKeys.VS.Normal <> 0 AndAlso FormKeys.VS.Late = 0 Then
                AppSwitcherContext.OpenApp(Me, "Microsoft Visual Studio")
            ElseIf FormKeys.VS.Normal = 0 AndAlso FormKeys.VS.Late <> 0 Then
                AppSwitcherContext.CloseApp(Me)
            End If
            FormKeys.VS.Late = FormKeys.VS.Normal

            FormKeys.SQL.Normal = CByte(data(4) And 8)
            If FormKeys.SQL.Normal <> 0 AndAlso FormKeys.SQL.Late = 0 Then
                AppSwitcherContext.OpenApp(Me, "Microsoft SQL Server Management Studio")
            ElseIf FormKeys.SQL.Normal = 0 AndAlso FormKeys.SQL.Late <> 0 Then
                AppSwitcherContext.CloseApp(Me)
            End If
            FormKeys.SQL.Late = FormKeys.SQL.Normal


            FormKeys.Pims.Normal = CByte(data(3) And 16)
            If FormKeys.Pims.Normal <> 0 AndAlso FormKeys.Pims.Late = 0 Then
                AppSwitcherContext.OpenApp(Me, "Pims")
                SetFlachOnKey(4, True)
            ElseIf FormKeys.Pims.Normal = 0 AndAlso FormKeys.Pims.Late <> 0 Then
                AppSwitcherContext.CloseApp(Me)
                SetFlachOnKey(4, False)
            End If
            FormKeys.Pims.Late = FormKeys.Pims.Normal

            FormKeys.Chrome.Normal = CByte(data(4) And 16)

            If FormKeys.Chrome.Normal <> 0 AndAlso FormKeys.Chrome.Late = 0 Then
                AppSwitcherContext.OpenApp(Me, "Chrome")

            ElseIf FormKeys.Chrome.Normal = 0 AndAlso FormKeys.Chrome.Late <> 0 Then
                AppSwitcherContext.CloseApp(Me)
            End If
            FormKeys.Chrome.Late = FormKeys.Chrome.Normal

            FormKeys.Outlock.Normal = CByte(data(3) And 32)
            If FormKeys.Outlock.Normal <> 0 AndAlso FormKeys.Outlock.Late = 0 Then
                AppSwitcherContext.OpenApp(Me, "Outlock")

                SetFlachOnKey(5, True, True)
            ElseIf FormKeys.Outlock.Normal = 0 AndAlso FormKeys.Outlock.Late <> 0 Then
                AppSwitcherContext.CloseApp(Me)
            End If
            FormKeys.Outlock.Late = FormKeys.Outlock.Normal


            FormKeys.Skype.Normal = CByte(data(4) And 32)
            If FormKeys.Skype.Normal <> 0 AndAlso FormKeys.Skype.Late = 0 Then

                SetFlachOnKey(13, False)
            End If
            FormKeys.Skype.Late = FormKeys.Skype.Normal

            FormKeys.PlayKey.Normal = CByte(data(11) And 8)
            If FormKeys.PlayKey.Normal <> 0 AndAlso FormKeys.PlayKey.Late = 0 Then
                Me.Invoke(DirectCast(Sub()
                                         SendKeys.Send("{F5}")
                                     End Sub, Action))
                SetFlachOnKey(26, True, True)
            ElseIf FormKeys.PlayKey.Normal = 0 AndAlso FormKeys.PlayKey.Late <> 0 Then
                SetFlachOnKey(26, False)

            End If
            FormKeys.PlayKey.Late = FormKeys.PlayKey.Normal


            FormKeys.StopKey.Normal = CByte(data(12) And 8)
            If FormKeys.StopKey.Normal <> 0 AndAlso FormKeys.StopKey.Late = 0 Then
                Me.Invoke(DirectCast(Sub()
                                         SendKeys.Send("+{F5}")
                                     End Sub, Action))
                SetFlachOnKey(26, True, True)
            ElseIf FormKeys.StopKey.Normal = 0 AndAlso FormKeys.StopKey.Late <> 0 Then
                SetFlachOnKey(26, False)

            End If
            FormKeys.StopKey.Late = FormKeys.StopKey.Normal
        End If

    End Sub

    Public Sub HandlePIEHidError(sourceDevices As PIEDevice, [error] As Integer) Implements PIEErrorHandler.HandlePIEHidError
        Throw New NotImplementedException()
    End Sub

#End Region

#Region "X-Keys Functions"


    Private Sub SetFlachOnKey(Index As Integer, pOn As Boolean, Optional pFlash As Boolean = False)
        If devices Is Nothing Then Return
        For i As Integer = 0 To devices(selecteddevice).WriteLength - 1
            wdata(i) = 0
        Next
        wdata(0) = 0
        wdata(1) = 181
        wdata(2) = Index

        If pOn Then
            wdata(3) = 1
            If pFlash Then
                wdata(3) = 2
            End If
        End If


        Dim result As Integer
        result = 404
        While (result = 404)
            result = devices(selecteddevice).WriteData(wdata)
        End While
    End Sub

    Private Sub Back1OnOff(pBool As Boolean)
        'Turns on or off ALL bank 1 BLs using current intensity
        If selecteddevice <> -1 Then
            Dim sl As Byte = 0

            If pBool = True Then
                sl = 255
            End If

            For i As Integer = 0 To devices(selecteddevice).WriteLength - 1
                wdata(i) = 0
            Next
            wdata(0) = 0
            wdata(1) = 182 'b6
            wdata(2) = 0 '0 for bank1, 1 for bank 2
            wdata(3) = sl

            Dim result As Integer
            result = 404
            While (result = 404)
                result = devices(selecteddevice).WriteData(wdata)
            End While
        End If
    End Sub

    Private Sub Setup()
        'setup devices for data and error callbacks
        If CboDevices.Count > 0 Then
            For i As Integer = 0 To CboDevices.Count - 1
                devices(cbotodevice(i)).SetDataCallback(Me)
                devices(cbotodevice(i)).SetErrorCallback(Me)
                devices(cbotodevice(i)).callNever = False
            Next
            selecteddevice = cbotodevice(0)
        End If
    End Sub


    Private Sub Init()
        'do this first to get the devices connected
        EnumerationSuccess = False
        selecteddevice = -1 'means no device is selected
        CboDevices.Clear()
        devices = PIEHid32Net.PIEDevice.EnumeratePIE()
        If devices.Length = 0 Then
            '   LblStatus.Text = "No Devices Found"
        Else
            Dim cbocount As Integer = 0
            For i As Integer = 0 To devices.Length - 1

                If devices(i).HidUsagePage = 12 And devices(i).WriteLength = 36 Then

                    Select Case devices(i).Pid
                        Case 1089
                            CboDevices.Add("XK-80 (" + devices(i).Pid.ToString + "=PID #1)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1090
                            CboDevices.Add("XK-80 (" + devices(i).Pid.ToString + "=PID #2)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1091
                            CboDevices.Add("XK-80 (" + devices(i).Pid.ToString + "=PID #3)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1250
                            CboDevices.Add("XK-80 (" + devices(i).Pid.ToString + "=PID #4)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1121
                            CboDevices.Add("XK-60 (" + devices(i).Pid.ToString + "=PID #1)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1122
                            CboDevices.Add("XK-60 (" + devices(i).Pid.ToString + "=PID #2)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1123
                            CboDevices.Add("XK-60 (" + devices(i).Pid.ToString + "=PID #3)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case 1254
                            CboDevices.Add("XK-60 (" + devices(i).Pid.ToString + "=PID #4)")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                        Case Else
                            CboDevices.Add("Unknown Device (" + devices(i).Pid.ToString + ")")
                            cbotodevice(cbocount) = i
                            cbocount = cbocount + 1
                    End Select
                    Dim result As Integer = devices(i).SetupInterface()
                    devices(i).suppressDuplicateReports = False ' ChkSuppress.Checked


                End If
            Next
        End If

        If CboDevices.Count > 0 Then
            selecteddevice = cbotodevice(0)

            ReDim wdata(devices(selecteddevice).WriteLength - 1) 'initialize length of write buffer

            EnumerationSuccess = True
            Me.Cursor = Cursors.Default
        End If
    End Sub
#End Region

    Private Function getKeys(pHotKey As String)
        Dim vReturn As New List(Of String)
        If Not pHotKey.Contains("+") Then
            vReturn.Add(pHotKey)
            Return vReturn
        End If

        Dim vHotKeys As String() = pHotKey.Split("+")

        Dim vHotKeys2 As New List(Of String)
        For Each value In vHotKeys
            vHotKeys2.Add(value.Replace("+", ""))
        Next

        Return vHotKeys2
    End Function

    Private Function IsHotKeyPressed(pHotKey As Nullable(Of KeyValuePair(Of String, String))) As Boolean
        Dim vError = True
        If pHotKey Is Nothing Then Return False

    End Function

    Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If Config.useXkey Then
            Return
        End If

        For Each app In RegestryContext.Apps
            Dim vHoyKey As Nullable(Of KeyValuePair(Of String, String)) = (From key In RegestryContext.HotKeys
                                                                           Where key.Key = app.Name).FirstOrDefault()
            If vHoyKey Is Nothing Then Return



        Next

        If e.Control AndAlso e.KeyCode = Keys.Q Then
            Me.Invoke(DirectCast(Sub()
                                     Form = New AppSwither("Microsoft Visual Studio")
                                     CloseForms()
                                     FormsPool.Add(Form)
                                     Form.Show()
                                     Form.Focus()

                                 End Sub, Action))
            SetFlachOnKey(3, True)
        ElseIf FormKeys.VS.Normal = 0 AndAlso FormKeys.VS.Late <> 0 Then
            Me.Invoke(DirectCast(Sub()
                                     Form.Close()
                                 End Sub, Action))
        End If
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RegestryContext.Init()
        RegestryContext.LoadData()
        Config.MainForm = Me
        Me.KeyPreview = True

        If Config.useXkey Then
            Me.Hide()
        End If

        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.BackColor = Color.Magenta
        Me.TransparencyKey = Color.Magenta
        NotifyIcon1.Visible = True

        Me.NotifyIcon1.Text = "AppSwitcher++"
        AddHandler NotifyIcon1.Click, AddressOf NotifyIcon1Clicked
        Dim vContextMenu As New ContextMenu()
        vContextMenu.MenuItems.Add("Config")
        Me.NotifyIcon1.ContextMenu = vContextMenu
        Me.NotifyIcon1.Visible = True
        Me.ShowInTaskbar = False
        If Not Config.useXkey Then
        Else
            selecteddevice = -1
            Init()
            Setup()
            Back1OnOff(False)
        End If

        RegestryContext.Regester(Me)

    End Sub

    Private Sub NotifyIcon1Clicked(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            If Config.ConfigForm Is Nothing Then
                Config.ConfigForm = New ConfigForm
            End If
            If Config.ConfigForm.Visible = False Then
                Config.ConfigForm.ClearForm()
                Config.ConfigForm.ShowDialog()
            End If

        End If

    End Sub

    Public Sub InitXKeys()
        selecteddevice = -1
        Init()
        Setup()
        Back1OnOff(False)
    End Sub


    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)
        Me.Hide()
    End Sub

    Private Sub NotifyIcon1_Click(sender As Object, e As EventArgs)
        Me.Show()
    End Sub

    Private Sub NotifyIcon1_BalloonTipClicked(sender As Object, e As EventArgs)
        Me.Show()
    End Sub
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H312 Then

            Dim key As Keys = DirectCast((CInt(m.LParam) >> 16) And &HFFFF, Keys)
            ' The key of the hotkey that was pressed.
            Dim modifier As KeyModifier = DirectCast(CInt(m.LParam) And &HFFFF, KeyModifier)
            ' The modifier of the hotkey that was pressed.
            Dim id As Integer = m.WParam.ToInt32()
            ' The id of the hotkey that was pressed.
            Dim vInput As New ShortcutInput
            Select Case modifier
                Case KeyModifier.Control
                    vInput.Control = True
                Case KeyModifier.Alt
                    vInput.Alt = True
                Case KeyModifier.Shift
                    vInput.Shift = True
            End Select
            vInput.Value2 = vInput.CharCodeFromKeys2(key)
            vInput.ClearDdnChars()
            Dim vCharCode As Integer = CType(vInput.Keys, Integer)

            Dim vKey = (From k In RegestryContext.HotKeys
                        Where k.Value.Trim = vCharCode.ToString.Trim
                        Select k).FirstOrDefault()

            AppSwitcherContext.OpenApp(Me, vKey.Key)
            AppSwitcherContext.Current = vCharCode
            AppSwitcherContext.CharKey = CType(key, Integer)

            Me.Activate()
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub AppSwither2_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If Config.useXkey Then Return

        If e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Right Then
            AppSwitcherContext.Form.AppSwither_KeyDown(sender, e)
        End If

    End Sub

    Private Sub Main_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        If Config.useXkey Then Return
        If e.KeyCode = AppSwitcherContext.CharKey Then
            AppSwitcherContext.CloseApp(Me)
        End If
    End Sub


End Class


