Public Class ConfigForm
    Public Sub ClearForm()
        BindingSource1.DataSource = New List(Of AppGridObject)
        DataGridView1.Refresh()
    End Sub

    Private Sub ConfigForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BindingSource1.DataSource = New List(Of AppGridObject)

        If Config.useXkey Then
            CheckBox1.Checked = True
        End If

        DataGridView1.Columns.Clear()
        Dim vNameCol As New DataGridViewTextBoxColumn With {.HeaderText = "Name", .Name = "Name", .DataPropertyName = "Name"}
        DataGridView1.Columns.Add(vNameCol)
        Dim vPathCol As New DataGridViewTextBoxColumn With {.HeaderText = "Path", .Name = "Path", .DataPropertyName = "Path", .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells}
        DataGridView1.Columns.Add(vPathCol)

        Dim vHotKey As New DataGridViewButtonColumn With {.HeaderText = "Change HotKey", .DataPropertyName = "HotKey"}
        DataGridView1.Columns.Add(vHotKey)

        Dim vListOfApps As New List(Of AppGridObject)
        For Each app In RegestryContext.Apps
            Dim vApp As New AppGridObject With {.Name = app.Name, .Path = app.Path}
            vListOfApps.Add(vApp)
        Next

        BindingSource1.DataSource = vListOfApps
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Config.useXkey = CheckBox1.Checked
        If CheckBox1.Checked Then
            Config.MainForm.InitXKeys()
        End If

        RegestryContext.Reset()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 2 Then
            Dim vHoyKeyDialog As New SpecifyShortcut(DataGridView1.CurrentRow.Cells(0).Value, Keys.A)
            If vHoyKeyDialog.ShowDialog() = DialogResult.OK Then
                If RegestryContext.HotKeys.Keys.Contains(DataGridView1.CurrentRow.Cells(0).Value.ToString) Then
                    RegestryContext.HotKeys.Remove(DataGridView1.CurrentRow.Cells(0).Value.ToString)
                End If
                Dim vHotKey As String = CType(vHoyKeyDialog.ShortcutInput1.Keys, Integer).ToString
                RegestryContext.HotKeys.Add(DataGridView1.CurrentRow.Cells(0).Value.ToString, vHotKey)

                RegestryContext.Reset()
                RegestryContext.Regester(Me)
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        'TODO: Save it to register her

        Dim vAppName As String = ""
        Dim vPath As String = ""

        If DataGridView1.Rows(e.RowIndex).Cells(0).Value IsNot Nothing Then
            vAppName = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString
        End If

        If DataGridView1.Rows(e.RowIndex).Cells(1).Value IsNot Nothing Then
            vPath = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString
        End If

        Dim vCurrent = RegestryContext.Apps.Where(Function(app) app.Name = vAppName).FirstOrDefault()
        If vCurrent IsNot Nothing Then
            RegestryContext.Apps.Remove(vCurrent)
        End If

        If Name = "" Then Return
        Dim vNew As New Program With {.Name = vAppName, .Path = vPath}
        RegestryContext.Apps.Add(vNew)
        RegestryContext.Reset()

    End Sub

    Private Sub DataGridView1_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
        Dim vAppName As String = ""
        If DataGridView1.Rows.Count = 0 Then Return
        If DataGridView1.Rows(e.Row.Index).Cells(0).Value IsNot Nothing Then
            vAppName = DataGridView1.Rows(e.Row.Index).Cells(0).Value.ToString
        End If

        Dim vCurrent = RegestryContext.Apps.Where(Function(app) app.Name = vAppName).FirstOrDefault()
        If vCurrent IsNot Nothing Then
            RegestryContext.Apps.Remove(vCurrent)
        End If

        RegestryContext.Reset()
        RegestryContext.Regester(Me)
    End Sub
End Class