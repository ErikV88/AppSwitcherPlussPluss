Module AppSwitcherContext

    Public Current As Integer
    Public CharKey As Integer
    Public Property ModiferKey As Keys
    Public Property CharKey2 As Keys

    Public Function getSynonymKey(pKey As Keys) As Keys
        If pKey = Keys.LButton Then Return Keys.Menu
        If pKey = Keys.Menu Then Return Keys.Alt
            Return pKey
    End Function

    Public Sub OpenApp(sender As Form, pAppName As String)
        If Config.AppSwither IsNot Nothing AndAlso Config.useXkey Then
            Config.AppSwither.Dispose()
        End If

        Dim vAppExist As Boolean = False
        For Each app In RegestryContext.Apps
            If app.Name = pAppName Then vAppExist = True
        Next
        If Not vAppExist Then Return
        If Current > 0 Then Return
        sender.Invoke(DirectCast(Sub()
                                     If Config.AppSwither IsNot Nothing AndAlso Config.useXkey Then
                                         Config.AppSwither.Dispose()
                                     End If
                                     Config.AppSwither = New AppSwither(pAppName)

                                     Config.AppSwither.Show()
                                     Config.AppSwither.Focus()

                                 End Sub, Action))

    End Sub

    Public Sub CloseApp(sender As Form)
        If Config.AppSwither Is Nothing Then Return
        sender.Invoke(DirectCast(Sub()
                                     Config.AppSwither.Close()
                                     Config.AppSwither.Dispose()
                                 End Sub, Action))
        Current = -1
    End Sub
End Module
