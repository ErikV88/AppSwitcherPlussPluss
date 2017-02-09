
' 
' 2004 Copyright Ashley van Gerven (ashley_van_gerven@hotmail.com)
' > Feel free to use/modify/delete this control
' > Feel free to remove this copyright notice (if you *need* to!)
' > Feel free to let me know of any cool things you do with this code (ie improvements etc)
' 
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Windows.Forms



Public Class ShortcutInput
    Inherits System.Windows.Forms.UserControl

#Region "Public Properties"
    Public Value2 As Integer

    Public Sub ClearDdnChars()
        DdnChars.Items.Clear()
    End Sub

    Public Property CharCode() As Byte
        Get
            If DdnChars.Items.Count = 0 Then Return CByte(Value2)
            Return Convert.ToByte(DdnChars.SelectedItem(1))
        End Get
        Set

            For Each item As Object In DdnChars.Items
                If item.ToString() = " " + ChrW(Value.ToString) Then
                    DdnChars.SelectedItem = item
                    Return
                End If
            Next
        End Set
    End Property




    Public ReadOnly Property Win32Modifiers() As Byte
            Get
                Dim toReturn As Byte = 0
                If CbxShift.Checked Then
                    toReturn += ModShift
                End If
                If CbxControl.Checked Then
                    toReturn += ModControl
                End If
                If CbxAlt.Checked Then
                    toReturn += ModAlt
                End If
                Return toReturn
            End Get
        End Property


        Public Property Keys() As Keys
            Get
                Dim k As Keys = CharCode
                If CbxShift.Checked Then
                    k = k Or Keys.Shift
                End If
                If CbxControl.Checked Then
                    k = k Or Keys.Control
                End If
                If CbxAlt.Checked Then
                    k = k Or Keys.Alt
                End If
                Return k
            End Get
            Set
                Dim k As Keys = Value
                If (CInt(k) And CInt(Keys.Shift)) = CInt(Keys.Shift) Then
                    Shift = True
                End If
                If (CInt(k) And CInt(Keys.Control)) = CInt(Keys.Control) Then
                    Control = True
                End If
                If (CInt(k) And CInt(Keys.Alt)) = CInt(Keys.Alt) Then
                    Alt = True
                End If

                CharCode = ShortcutInput.CharCodeFromKeys(k)
            End Set
        End Property


        Public Property Shift() As Boolean
            Get
                Return CbxShift.Checked
            End Get
            Set
                CbxShift.Checked = Value
            End Set
        End Property


        Public Property Control() As Boolean
            Get
                Return CbxControl.Checked
            End Get
            Set
                CbxControl.Checked = Value
            End Set
        End Property


        Public Property Alt() As Boolean
            Get
                Return CbxAlt.Checked
            End Get
            Set
                CbxAlt.Checked = Value
            End Set
        End Property


        Public MinModifiers As Byte = 0


        Public ReadOnly Property IsValid() As Boolean
            Get
                Dim ModCount As Byte = 0
                ModCount += CByte(If((Shift), 1, 0))
                ModCount += CByte(If((Control), 1, 0))
                ModCount += CByte(If((Alt), 1, 0))
                If ModCount < MinModifiers Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property
#End Region


        Private Const ModAlt As Byte = 1, ModControl As Byte = 2, ModShift As Byte = 4, ModWin As Byte = 8
        Private CbxShift As System.Windows.Forms.CheckBox
        Private CbxControl As System.Windows.Forms.CheckBox
        Private CbxAlt As System.Windows.Forms.CheckBox
        Private DdnChars As System.Windows.Forms.ComboBox
        ''' <summary> 
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.Container = Nothing


    Public Sub New()
        ' This call is required by the Windows.Forms Form Designer.
        InitializeComponent()

        For i As Integer = 65 To 90
            DdnChars.Items.Add(" " + ChrW(i.ToString))
        Next

        For i As Integer = 48 To 57
            DdnChars.Items.Add(" " + ChrW(i.ToString))
        Next

        DdnChars.SelectedIndex = 0
    End Sub


    Public Sub AddMore()
        For i As Integer = 65 To 90
            DdnChars.Items.Add(" " + CChar(i.ToString))
        Next

        For i As Integer = 48 To 57
            DdnChars.Items.Add(" " + CChar(i.ToString))
        Next

        DdnChars.SelectedIndex = 0
    End Sub


    ''' <summary>
    ''' Calculates the Win32 Modifiers total for a Keys enum
    ''' </summary>
    ''' <param name="k">An instance of the Keys enumaration</param>
    ''' <returns>The Win32 Modifiers total as required by RegisterHotKey</returns>
    Public Shared Function Win32ModifiersFromKeys(k As Keys) As Byte
            Dim total As Byte = 0

            If (CInt(k) And CInt(Keys.Shift)) = CInt(Keys.Shift) Then
                total += ModShift
            End If
            If (CInt(k) And CInt(Keys.Control)) = CInt(Keys.Control) Then
                total += ModControl
            End If
            If (CInt(k) And CInt(Keys.Alt)) = CInt(Keys.Alt) Then
                total += ModAlt
            End If
            If (CInt(k) And CInt(Keys.LWin)) = CInt(Keys.LWin) Then
                total += ModWin
            End If

            Return total
        End Function


    ''' <summary>
    ''' Calculates the character code of alphanumeric key of the Keys enum instance
    ''' </summary>
    ''' <param name="k">An instance of the Keys enumaration</param>
    ''' <returns>The character code of the alphanumeric key</returns>
    Public Function CharCodeFromKeys2(k As Keys) As Byte
        Dim charCode As Byte = 0
        If (k.ToString().Length = 1) OrElse ((k.ToString().Length > 2) AndAlso (k.ToString()(1) = ","c)) Then
            charCode = Microsoft.VisualBasic.AscW(k.ToString()(0))
        ElseIf (k.ToString().Length > 3) AndAlso (k.ToString()(0) = "D"c) AndAlso (k.ToString()(2) = ","c) Then
            charCode = Microsoft.VisualBasic.AscW(k.ToString()(1))
        End If
        Return charCode
    End Function


    Public Shared Function CharCodeFromKeys(k As Keys) As Byte
        Dim charCode As Byte = 0
        If (k.ToString().Length = 1) OrElse ((k.ToString().Length > 2) AndAlso (k.ToString()(1) = ","c)) Then
            charCode = Microsoft.VisualBasic.AscW(k.ToString()(0))
        ElseIf (k.ToString().Length > 3) AndAlso (k.ToString()(0) = "D"c) AndAlso (k.ToString()(2) = ","c) Then
            charCode = Microsoft.VisualBasic.AscW(k.ToString()(1))
        End If
        Return charCode
    End Function


    ''' <summary> 
    ''' Clean up any resources being used.
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                If components IsNot Nothing Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub


#Region "Component Designer generated code"
        ''' <summary> 
        ''' Required method for Designer support - do not modify 
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
        Me.CbxShift = New System.Windows.Forms.CheckBox()
        Me.CbxControl = New System.Windows.Forms.CheckBox()
        Me.CbxAlt = New System.Windows.Forms.CheckBox()
        Me.DdnChars = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'CbxShift
        '
        Me.CbxShift.Location = New System.Drawing.Point(8, 8)
        Me.CbxShift.Name = "CbxShift"
        Me.CbxShift.Size = New System.Drawing.Size(78, 24)
        Me.CbxShift.TabIndex = 0
        Me.CbxShift.Text = "Shift"
        '
        'CbxControl
        '
        Me.CbxControl.Location = New System.Drawing.Point(92, 8)
        Me.CbxControl.Name = "CbxControl"
        Me.CbxControl.Size = New System.Drawing.Size(101, 24)
        Me.CbxControl.TabIndex = 1
        Me.CbxControl.Text = "Control"
        '
        'CbxAlt
        '
        Me.CbxAlt.Location = New System.Drawing.Point(190, 8)
        Me.CbxAlt.Name = "CbxAlt"
        Me.CbxAlt.Size = New System.Drawing.Size(101, 24)
        Me.CbxAlt.TabIndex = 2
        Me.CbxAlt.Text = "Alt"
        '
        'DdnChars
        '
        Me.DdnChars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.DdnChars.Location = New System.Drawing.Point(257, 8)
        Me.DdnChars.Name = "DdnChars"
        Me.DdnChars.Size = New System.Drawing.Size(138, 28)
        Me.DdnChars.TabIndex = 4
        '
        'ShortcutInput
        '
        Me.Controls.Add(Me.DdnChars)
        Me.Controls.Add(Me.CbxAlt)
        Me.Controls.Add(Me.CbxControl)
        Me.Controls.Add(Me.CbxShift)
        Me.Name = "ShortcutInput"
        Me.Size = New System.Drawing.Size(478, 48)
        Me.ResumeLayout(False)

    End Sub
#End Region

End Class
