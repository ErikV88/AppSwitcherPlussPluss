Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms


Public Class SpecifyShortcut
    Inherits System.Windows.Forms.Form
    Private WithEvents BtnOK As System.Windows.Forms.Button
    Private WithEvents BtnCancel As System.Windows.Forms.Button
    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.Container = Nothing
    Public ShortcutInput1 As ShortcutInput
    Private AppName As String

    Public Sub New(pAppName As String, k As Keys)
        AppName = pAppName
        ' Required for Windows Form Designer support

        InitializeComponent()

        Me.Text = String.Format("HotKey for {0}", pAppName)

        If k <> Keys.None Then
            ShortcutInput1.Keys = k
        End If
    End Sub


    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not Config.useXkey Then Return
        If disposing Then
            If components IsNot Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub


#Region "Windows Form Designer generated code"
    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.ShortcutInput1 = New ShortcutInput()
        Me.SuspendLayout()
        '
        'BtnOK
        '
        Me.BtnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnOK.Location = New System.Drawing.Point(77, 129)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(120, 33)
        Me.BtnOK.TabIndex = 0
        Me.BtnOK.Text = "&OK"
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnCancel.Location = New System.Drawing.Point(205, 129)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(120, 33)
        Me.BtnCancel.TabIndex = 1
        Me.BtnCancel.Text = "&Cancel"
        '
        'ShortcutInput1
        '
        Me.ShortcutInput1.Alt = False
        Me.ShortcutInput1.Control = False
        Me.ShortcutInput1.Location = New System.Drawing.Point(13, 23)
        Me.ShortcutInput1.Name = "ShortcutInput1"
        Me.ShortcutInput1.Shift = False
        Me.ShortcutInput1.Size = New System.Drawing.Size(437, 71)
        Me.ShortcutInput1.TabIndex = 2
        '
        'FrmSpecifyShortcut
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(462, 197)
        Me.Controls.Add(Me.ShortcutInput1)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmSpecifyShortcut"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Specify Shortcut"
        Me.ResumeLayout(False)

    End Sub
#End Region


    Private Sub BtnOK_Click(sender As Object, e As System.EventArgs) Handles BtnOK.Click
        If Not ShortcutInput1.IsValid Then
            '        MessageBox.Show("Please check at least " + MinModifiers + " of the checkboxes (ie Modifier keys)   ", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return
        End If

        DialogResult = DialogResult.OK
        Close()
    End Sub

End Class
