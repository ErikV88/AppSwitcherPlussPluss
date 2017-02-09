﻿Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Module RegestryContext

    <DllImport("user32.dll")>
    Public Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As Integer, vlc As Integer) As Boolean
    End Function

    Public Property Apps As New List(Of Program)
    Public Property HotKeys As New Dictionary(Of String, String)
    Public Property useXKey As Integer = 0

    Public Sub Init()
        If Registry.CurrentUser.GetValue("Software\\AppSwitcher\\") IsNot Nothing Then
            Return
        End If

        Reset()
    End Sub

    Public Sub Reset()
        If Registry.CurrentUser.GetValue("Software\\AppSwitcher\\") IsNot Nothing Then
            Registry.CurrentUser.DeleteSubKey("Software\\AppSwitcher\\")
        End If

        Registry.CurrentUser.CreateSubKey("Software\\AppSwitcher\\")
        Registry.CurrentUser.CreateSubKey("Software\\AppSwitcher\\Apps")
        Registry.CurrentUser.CreateSubKey("Software\\AppSwitcher\\HotKeys")

        Try
            Registry.CurrentUser.OpenSubKey("Software\\AppSwitcher\\", True).SetValue("UseXKeys", useXKey.ToString, Microsoft.Win32.RegistryValueKind.String)

        Catch ex As Exception

        End Try

        For Each a In Apps
            Registry.CurrentUser.OpenSubKey("Software\\AppSwitcher\\Apps\\", True).SetValue(a.Name, a.Path, Microsoft.Win32.RegistryValueKind.String)
        Next

        For Each h In HotKeys
            Registry.CurrentUser.OpenSubKey("Software\\AppSwitcher\\HotKeys\\", True).SetValue(h.Key, h.Value, Microsoft.Win32.RegistryValueKind.String)
        Next
    End Sub

    Public Sub Regester(sender As Form)
        For Each key In RegestryContext.HotKeys
            Dim val As Integer = -1
            Try
                If key.Value.Trim() <> "" Then
                    val = Int32.Parse(key.Value.Trim())
                End If
            Catch
            End Try
            If val <> -1 Then
                Dim k As Keys = DirectCast(val, Keys)
                Dim success As Boolean = RegisterHotKey(sender.Handle, sender.[GetType]().GetHashCode(), ShortcutInput.Win32ModifiersFromKeys(k), ShortcutInput.CharCodeFromKeys(k))
                If success Then
                    '  TxtKeyEnumValue.Text = val.ToString()
                Else
                    'MessageBox.Show("Could not register Hotkey - there is probably a conflict.  ", "", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End If
            End If
        Next
    End Sub


    Public Sub LoadData()
        Apps.Clear()
        HotKeys.Clear()
        If Registry.CurrentUser.GetValue("Software\\AppSwitcher") IsNot Nothing Then

            Return
        End If

        Dim vUseXKey As Integer = 0
        useXKey = Integer.TryParse(Registry.CurrentUser.OpenSubKey("Software\\AppSwitcher").GetValue("UseXKey"), vUseXKey)

        Try
            For Each key In GetKeyValues("Software\\AppSwitcher\\Apps")
                Apps.Add(New Program With {.Name = key.Key, .Path = key.Value})
            Next

            For Each key In GetKeyValues("Software\\AppSwitcher\\HotKeys")
                HotKeys.Add(key.Key, key.Value)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Function GetKeyValues(keyName As String) As Dictionary(Of String, String)
        Dim result = New Dictionary(Of String, String)

        Using key = Registry.CurrentUser.OpenSubKey(keyName)
            For Each valueName In key.GetValueNames()
                Dim val = key.GetValue(valueName)

                result.Add(valueName, If(val = "", " ", val))
            Next
        End Using
        Return result
    End Function

End Module