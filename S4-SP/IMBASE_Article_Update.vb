Imports System.IO
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Type
Imports System.Activator
Public Class IMBASE_Article_Update
    Private imb As ImBase.ImbaseApplication
    Public _invApp As Inventor.Application ' = GetObject(, "Inventor.Application") 
    Dim started As Boolean = False
    Private BuferFolder As String = "C:\Users\aidarhanov.n.VEZA-SPB\Desktop\test 1\буффер" ' "C:\Users\aidarhanov.n.VEZA-SPB\Desktop\tmp"
    Private ParentFolder As String = "C:\Users\aidarhanov.n.VEZA-SPB\Desktop\test 1\поиск" ' "\\intermech\IM\CAD\Inventor\Models2018\Library"
    Private Start_Date As Date = "01.03.2019"
    Private Finish_Date As Date = Now
    Private prop_puth_name As String = "prop_puth"
    Public razdelitel As String = "@%@"
    Public Sub New()
        ' Обязательный вызов, нужен для дизайнера
        InitializeComponent()
        ' здесь размещается любой инициализирующий код.
        Try
            _invApp = Marshal.GetActiveObject("Inventor.Application")


        Catch ex As Exception
            Try
                Dim invAppType As Type =
                  GetTypeFromProgID("Inventor.Application")

                _invApp = CreateInstance(invAppType)
                _invApp.Visible = True

                'Примечание:
                'Данная логическая переменная служит индикатором, 
                'был ли данный сеанс Inventor создан нашей программой 
                'и требуется ли принудительное завершение сеанса
                'по окончании работы нашего приложения.
                started = True

            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Не удалось ни найти, ни создать сеанс Inventor", "DXF-Auto")
            End Try
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        MonthCalendar1.Visible = 1
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        MonthCalendar2.Visible = 1
    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        Start_Date = MonthCalendar1.SelectionStart
        TextBox1.Text = Start_Date
        MonthCalendar1.Visible = False
    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        Finish_Date = MonthCalendar2.SelectionStart
        TextBox2.Text = Finish_Date
        MonthCalendar2.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Буферная пака
        BuferFolder = GetSelectFolder()
        TextBox4.Text = BuferFolder
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Папка для поиска
        ParentFolder = GetSelectFolder()
        TextBox3.Text = ParentFolder
    End Sub
    Function GetSelectFolder()
        'возвращает выбранную папку
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Return FolderBrowserDialog1.SelectedPath
        End If
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'выйти
        Me.Close()
    End Sub
    Function chek_add_param() As Boolean
        'проверка заполеннности всех полей. Если все ОК - 1, если НЕТ - 0
        If TextBox1.Text = "" Then
            MsgBox("Не заполнена дата Начала отсчета")
            Return 0
        ElseIf TextBox2.Text = "" Then
            MsgBox("Не заполнена дата Конец отсчета")
            Return 0
        ElseIf TextBox3.Text = "" Then
            MsgBox("Не выбрана Папка поиска файлов")
            Return 0
        ElseIf TextBox4.Text = "" Then
            MsgBox("Не выбрана Буферная папка поиска файлов")
            Return 0
        Else
            Return 1
        End If
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'начать
        If chek_add_param() = 0 Then
            Exit Sub
        End If
        imb = CreateObject("ImBase.ImbaseApplication")

        Dim file_Array() As String = Directory.GetFiles(ParentFolder, "*.ipt", System.IO.SearchOption.AllDirectories)
        For Each tmp_file As String In file_Array
            Dim FInfo As FileInfo = New FileInfo(tmp_file)
            Dim LWrTime As Date = FInfo.LastWriteTime.ToShortDateString
            If (CDate(LWrTime) < CDate(Finish_Date)) And (CDate(Start_Date) < CDate(LWrTime)) Then
                Chanche_ContentCenter(tmp_file)
            End If
        Next
    End Sub
    Sub Chanche_ContentCenter(FPuth As String)

        Dim odc_temp As PartDocument = _invApp.Documents.Open(FPuth, False)
        Dim custom_propset_tmp As PropertySet = odc_temp.PropertySets.Item(4)

        Dim fname As String = System.IO.Path.GetFileName(odc_temp.FullFileName)

        'odc_temp.SaveAs(BuferFolder & "\" & fname, False)
        Try
            check_IMBASE_Key(odc_temp)
            odc_temp.Save()
        Catch ex As Exception

        End Try
        'odc_temp.SaveAs(FPuth, True)
        odc_temp.Close()
        ListBox1.Items.Add(FPuth & razdelitel & BuferFolder & "\" & fname)
        'Dim fileExists As Boolean = System.IO.File.Exists(FPuth)
        'If fileExists Then
        '    System.IO.File.Delete(FPuth)
        'End If
        'System.IO.File.Move(BuferFolder & "\" & fname, FPuth)
    End Sub

    Function check_IMBASE_Key(oDoc As Document)
        Dim MatNAme As String
        Dim matID As String
        Dim propset As PropertySets = oDoc.PropertySets

        MatNAme = propset.Item(4).Item("Наименование").Value
        matID = propset.Item(4).Item("Ключ IMBASE").Value
        Dim IBKey As ImBase.ImbaseKey = imb.Utilites.ExploreKey(matID)
        Dim iFN As String
        iFN = IBCatParser2(IBKey.TableRecStr, "Полное обозначение")
        propset.Item(4).Item("Наименование").Value = iFN & "6666"
        propset.Item(3).Item("Description").Value = iFN & "6666"
        oDoc.Update()
    End Function

    Function IBCatParser2(Fields As String, FieldName As String)
        Dim res As String
        Dim arr() As String
        arr = Split(Fields, vbCrLf)
        Dim s As String
        For i = 0 To UBound(arr) - 1
            s = arr(i)
            If InStr(s.ToUpper, FieldName.ToUpper) <> 0 Then
                res = Trim(Replace(Replace(Replace(s, FieldName, ""), FieldName.ToUpper, ""), "=", ""))
                Exit For
            End If
        Next
        Return res
    End Function

    Private Sub IMBASE_Article_Update_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Start_Date
        TextBox2.Text = Finish_Date
        TextBox3.Text = ParentFolder
        TextBox4.Text = BuferFolder
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        imb = CreateObject("ImBase.ImbaseApplication")
        Chanche_ContentCenter("\\intermech\IM\CAD\Inventor\Models2018\Library\Винт B.М3-6gx20.36.016 ГОСТ 1491-80.ipt")
        'If ListBox1.Items.Count = 0 Then Exit Sub
        'Dim tmp_line As String
        'Dim tmp_line_Arr() As String
        'For Each tmp_line In ListBox1.Items
        '    '0-файл оригинал, 2-буферный файл
        '    tmp_line_Arr = tmp_line.Split(razdelitel)

        '    Dim fileExists As Boolean = System.IO.File.Exists(tmp_line_Arr(0))
        '    If fileExists Then
        '        System.IO.File.Delete(tmp_line_Arr(0))
        '    End If
        '    System.IO.File.Move(tmp_line_Arr(2), tmp_line_Arr(0))
        'Next
    End Sub
End Class