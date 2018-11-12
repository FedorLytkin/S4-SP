Module Reester_edit_mod
    Public app_name = Application.ProductName
    Public tempform As Form
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
    Function GetReestrKeyExist(RegName As String) As Boolean
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & app_name & "\\0.9")
        If regVersion Is Nothing Then
            Return False
        Else
            If regVersion.GetValue(RegName) Is Nothing Then Return False Else Return True
        End If
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
    Public Function chek_reg_exist(regName As String) As Boolean
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & app_name & "\\0.9", True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                 "Software\\Microsoft\\" & app_name & "\\0.9")
        End If
        Dim regValue As String
        If regVersion IsNot Nothing Then
            regValue = regVersion.GetValue(regName)
            regVersion.Close()
        End If
        If regValue IsNot Nothing Then
            chek_reg_exist = True  ' параметра в регистре существует
        Else
            chek_reg_exist = False  ' пересетра в реестре НЕ существует
        End If
    End Function
    Public Function chek_reg_exist_with_appname(application_name As String, regName As String) As Boolean
        Dim regVersion = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  "Software\\Microsoft\\" & app_name & "\\0.9", True)
        If regVersion Is Nothing Then
            ' Key doesn't exist; create it.
            regVersion = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                 "Software\\Microsoft\\" & app_name & "\\0.9")
        End If
        Dim regValue As String
        If regVersion IsNot Nothing Then
            regValue = regVersion.GetValue(regName)
            regVersion.Close()
        End If
        If regValue IsNot Nothing Then
            chek_reg_exist_with_appname = True  ' параметра в регистре существует
        Else
            chek_reg_exist_with_appname = False  ' пересетра в реестре НЕ существует
        End If
    End Function

    Public Sub C()
        Try
            Dim file As System.IO.StreamReader = New System.IO.StreamReader("\\INTERMECH\im\Search\Search4.ini")

            Dim ld As String

            Dim counter As Integer = 0
            Dim line As String = " "
            While line IsNot Nothing
                line = file.ReadLine()
                If UCase(line).IndexOf(UCase("ldFiltr")) > 0 Then
                    ld = line.Substring(line.LastIndexOf("=") + 1).Trim(" ")
                    dCompare(ld)
                    Exit While
                End If
                counter += 1
            End While
            file.Close()
            If ld Is Nothing Or ld = "" Then Block_Controls(False)
        Catch ex As Exception
            Block_Controls(False)
        End Try
    End Sub

    Sub dCompare(Ldate As Date)
        Try
            If Ldate = "" Then Block_Controls(False) : Exit Sub
            Dim td As Date = Now.ToShortDateString
            If td > Ldate Then
                Block_Controls(False)
            Else
                Block_Controls(True)
            End If
        Catch ex As Exception
            Block_Controls(False)
        End Try
    End Sub




















    '\\модуль для блокировки программы в случае истечения лицензии
    'перед запуском функций необходимо присвоить значение переменной tempform As Form

    'эта процедура запускается вместе с программой и проверяет срок лицензии.
    'в случае истечения срока - все элементы приложения блокируются
    Public Sub Chek_licenc()
        Try
            Dim now_date As Date = Now.ToShortDateString
            Dim birthday As Date
            If chek_reg_exist("birthday") = False Then
                set_reesrt_value("birthday", now_date)
                birthday = get_reesrt_value("last_date", now_date)
            Else
                birthday = get_reesrt_value("last_date", now_date)
            End If

            Dim last_day As Date
            If chek_reg_exist("last_day") = False Then
                set_reesrt_value("last_day", birthday.AddMonths(1))
                last_day = get_reesrt_value("last_day", now_date)
            Else
                last_day = get_reesrt_value("last_day", now_date)
            End If


            If now_date > last_day Then
                'Me.Enabled = False
                Block_Controls(False)
            Else
                Block_Controls(True)
                'Me.Enabled = True
            End If
        Catch ex As Exception
            MsgBox("Ошибка! Обратитесь к разработчику")
            Block_Controls(False)
            'Me.Enabled = False
        End Try
    End Sub
    Public Sub Block_Controls(block_result As Boolean)
        Try
            For i As Integer = 0 To tempform.Controls.Count - 1
                tempform.Controls.Item(i).Enabled = block_result
            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Sub licCFG()
        Try
            Dim reques As String
            reques = InputBox("Введите Код запроса")
            Select Case reques
                Case ",juexfh" 'богучар на 1 месяц
                    lic_date_Update(1)
                Case "jcfrfhjdrf" ' осакаровка на 3 месяца
                    lic_date_Update(3)
                Case "cjhnbhjdrf" ' сортировка на 12 месяца
                    lic_date_Update(12)
                Case "frshpfvfy"  ' акырзаман обнуляет дату до сегодняшнего
                    lic_date_Update(0)
                Case vbNullString  ' выход из инпутбокса
                    Exit Sub
                Case Else
                    licCFG()
            End Select
        Catch ex As Exception

        End Try
    End Sub
    Public Sub lic_date_Update(Work_month_Count As Integer)
        Try
            Dim now_date As Date = Now.ToShortDateString
            Dim last_day As Date = get_reesrt_value("last_day", now_date)
            set_reesrt_value("last_day", now_date.AddMonths(Work_month_Count))
            last_day = get_reesrt_value("last_day", now_date)
            Chek_licenc()
        Catch ex As Exception

        End Try
    End Sub

End Module
