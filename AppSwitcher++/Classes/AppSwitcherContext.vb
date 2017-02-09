Module AppSwitcherContext
    Public Form As AppSwither

    Public Current As Integer
    Public CharKey As Integer

    Public Sub OpenApp(sender As Form, pAppName As String)
        Dim vAppExist As Boolean = False
        For Each app In RegestryContext.Apps
            If app.Name = pAppName Then vAppExist = True
        Next
        If Not vAppExist Then Return
        If Current > 0 Then Return
        sender.Invoke(DirectCast(Sub()
                                     Form = New AppSwither(pAppName)
                                     Form.Show()
                                     Form.Focus()

                                 End Sub, Action))

    End Sub

    Public Sub CloseApp(sender As Form)
        If Form Is Nothing Then Return
        sender.Invoke(DirectCast(Sub()
                                     Form.Close()
                                     Form.Dispose()
                                 End Sub, Action))
        Current = -1
    End Sub
End Module
