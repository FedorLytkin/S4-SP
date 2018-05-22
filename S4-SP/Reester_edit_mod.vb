Module Reester_edit_mod
    Public app_name = Application.ProductName
    Function get_reesrt_value(RegName As String, defolt_value As String)
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & app_name & "\\0.9", True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                 "Software\\Microsoft\\" & app_name & "\\0.9")
        End If
        Dim regValue As String
        If regVersion IsNot Nothing Then
            regValue = regVersion.GetValue(RegName, defolt_value)
            regVersion.Close()
        End If
        Return regValue
    End Function
    Sub set_reesrt_value(regName As String, regValue As String)
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & app_name & "\\0.9", True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                 "Software\\Microsoft\\" & app_name & "\\0.9")
        End If
        If regVersion IsNot Nothing Then
            regVersion.SetValue(regName, regValue)
            regVersion.Close()
        End If
    End Sub
    Function get_reesrt_value_with_appname(application_name As String, RegName As String, defolt_value As String)
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & application_name, True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                 "Software\\Microsoft\\" & application_name)
        End If
        Dim regValue As String
        If regVersion IsNot Nothing Then
            regValue = regVersion.GetValue(RegName, defolt_value)
            regVersion.Close()
        End If
        Return regValue
    End Function
    Sub set_reesrt_value_with_appname(application_name As String, regName As String, regValue As String)
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & application_name, True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                 "Software\\Microsoft\\" & application_name)
        End If
        If regVersion IsNot Nothing Then
            regVersion.SetValue(regName, regValue)
            regVersion.Close()
        End If
    End Sub
End Module
