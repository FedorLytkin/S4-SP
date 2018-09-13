Imports System.Windows.Forms
Public Class Art_property
    Public ArtID As Integer
    Public DocID As Integer
    Public a_oboz As TextBox = Me.TextBox1
    Public a_naim As TextBox = Me.TextBox2
    Public a_mater As TextBox = Me.TextBox3
    Public a_mass As TextBox = Me.TextBox4
    Public a_tolsh As TextBox = Me.TextBox5
    Public a_ArtID As TextBox = Me.TextBox8

    Public d_oboz_TB As TextBox = Me.TextBox10
    Public d_naim_TB As TextBox = Me.TextBox9
    Public d_format_CB As ComboBox = Me.ComboBox1
    Public d_Izm_CB As ComboBox = Me.ComboBox2
    Public d_DocID_TB As TextBox = Me.TextBox11

    Public S_Proj_AID_CB As ComboBox = Me.ComboBox5
    Public S_Format_CB As ComboBox = Me.ComboBox4
    Public S_Count_TB As TextBox = Me.TextBox6
    Public S_Positio_TB As TextBox = Me.TextBox12

    Public List_Format_Array As String = "БЧ; A0; A1; A2; A3; A4; A2x3; A3x3; A4x3; A4,A3; A1,A3; A2,A3; A1,A2; A1,A3x3; A2x3,A3"
    Dim str_ArdID As String = "ArtID: "
    Dim separator As String = " - "

    Private Sub Art_property_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txt_box_clear()
        List_Format_Array = get_reesrt_value("List_Format_Array", List_Format_Array)
        a_oboz = Me.TextBox1
        a_naim = Me.TextBox2
        a_mater = Me.TextBox3
        a_mass = Me.TextBox4
        a_tolsh = Me.TextBox5
        a_ArtID = Me.TextBox8

        d_oboz_TB = Me.TextBox10
        d_naim_TB = Me.TextBox9
        d_format_CB = Me.ComboBox1
        d_Izm_CB = Me.ComboBox2
        d_DocID_TB = Me.TextBox11

        S_Proj_AID_CB = Me.ComboBox5
        S_Count_TB = Me.TextBox6
        S_Positio_TB = Me.TextBox12

        Dim format_L As String
        Dim izm As String
        With Form1.ITS4App
            .OpenArticle(ArtID)
            Dim oboz As String = .GetFieldValue_Articles("Обозначение")
            Dim naim As String = .GetFieldValue_Articles("Наименование")
            Dim material As String = .GetFieldValue_Articles("Материал")
            Dim mass As String = .GetFieldValue_Articles("Масса")
            Dim tolshina As String = .GetFieldValue_Articles("Толщина/Диаметр")
            Me.Text = oboz & " " & naim
            a_oboz.Text = oboz
            a_naim.Text = naim
            a_mater.Text = material
            a_mass.Text = mass
            a_tolsh.Text = tolshina
            a_ArtID.Text = ArtID

            DocID = .GetDocID_ByArtID(ArtID)

            Button2.Enabled = False 'кнопка СОХРАНИТЬ не активна пока не нажата кнопка ИЗМЕНИТЬ
            Button3.Enabled = True
            If DocID > 0 Then
                .OpenDocument(DocID)
                Dim d_oboz As String = .GetFieldValue("Обозначение")
                Dim d_naim As String = .GetFieldValue("Наименование")
                Dim d_Version_ID As Integer = .GetFieldValue("Version_ID")
                format_L = .GetFieldValue("Формат")
                izm = .GetFieldValue("Изменение")
                .OpenDocVersion(DocID, d_Version_ID)
                Dim OTD_STATUS As String = .GetFieldValue_DocVersion("OTD_STATUS")
                .CloseDocument()
                Label8.Visible = False
                d_oboz_TB.Text = d_oboz
                d_naim_TB.Text = d_naim
                d_DocID_TB.Text = DocID
                ToolStripButton1.Enabled = True
                ToolStripButton2.Enabled = True

                If OTD_STATUS <> "" Or OTD_STATUS IsNot Nothing Then
                    Select Case OTD_STATUS
                        Case "0"
                            OTD_STATUS = "Не регистрирован"
                        Case "1"
                            OTD_STATUS = "Зарегистрирован"
                        Case "2"
                            OTD_STATUS = "Аннулирован"
                    End Select
                    ToolStripTextBox1.Text = "Статус ОТД: " & OTD_STATUS
                End If
            Else
                    Label8.Visible = True
                ToolStripButton1.Enabled = False
                ToolStripButton2.Enabled = False
            End If
            .CloseArticle()
        End With
        combobox_add(d_format_CB)
        combobox_add_Numeric(d_Izm_CB, 10)
        If format_L IsNot Nothing Or format_L <> "" Then
            d_format_CB.Items.Add(format_L)
            d_format_CB.Text = format_L
        End If
        If izm IsNot Nothing Or izm <> "" Then
            d_Izm_CB.Items.Add(izm)
            d_Izm_CB.Text = izm
        End If
        If Convert.ToBoolean(get_reesrt_value("ParamSostavVisible", False.ToString)) Then
            S_Proj_AID_CB_Items()
            Me.GroupBox3.Enabled = True
            Me.GroupBox3.Visible = True
            Me.Size = New Size(363, 494)
        Else
            Me.GroupBox3.Enabled = False
            Me.GroupBox3.Visible = False
            Me.Size = New Size(363, 361)
        End If
        read_only(True)

    End Sub
    Sub txt_box_clear()
        Dim allTextBoxValues As String = ""
        Dim c As Control
        Dim childc As Control
        For Each c In Me.Controls
            For Each childc In c.Controls
                If TypeOf childc Is TextBox Then
                    childc.Text = allTextBoxValues
                    'allTextBoxValues &= CType(childc, TextBox).Text & ","
                End If
            Next
        Next
    End Sub
    Sub combobox_add(CB As ComboBox)
        CB.Items.Clear()
        For Each txt As String In List_Format_Array.Split(";")
            txt = txt.Trim(" ")
            CB.Items.Add(txt)
        Next
    End Sub
    Sub S_Proj_AID_CB_Items()
        S_Proj_AID_CB.Items.Clear()
        Dim prj_id As String
        Try
            With Form1.ITS4App
                .OpenQuery("select * from pc where part_aid =" & ArtID)
                .QueryGoFirst()
                For i As Integer = 1 To .QueryRecordCount
                    prj_id = .QueryFieldByName("proj_aid")
                    .OpenArticle(prj_id)
                    Dim prj_oboz As String = .GetFieldValue_Articles("Обозначение")
                    Dim prj_naim As String = .GetFieldValue_Articles("Наименование")
                    S_Proj_AID_CB.Items.Add(str_ArdID & prj_id & separator & prj_oboz & separator & prj_naim)
                    .QueryGoNext()
                Next
                .CloseQuery()
            End With
            S_Proj_AID_CB.Text = S_Proj_AID_CB.Items.Item(0)
        Catch ex As Exception
            MsgBox("Ошибка при отправке SQL-запроса" + vbNewLine & ex.Message)
        End Try

        'Param_Sostav(prj_id)
    End Sub
    Public Sub Param_Sostav(prj_id As String)
        If prj_id IsNot Nothing Then
            With Form1.ITS4App
                Dim proj_link_ID As String
                Dim Pj_ArtID As Integer = .GetArtID_ByDesignation(prj_id)
                Try
                    .OpenQuery("select * from pc where part_aid = " & ArtID & " and proj_aid = " & Pj_ArtID)
                    proj_link_ID = .QueryFieldByName("PRJLINK_ID")
                    .CloseQuery()
                Catch ex As Exception
                    MsgBox("Ошибка при отправке SQL-запроса" + vbNewLine & ex.Message)
                End Try
                .OpenBOMItem(proj_link_ID)
                Try
                    Dim s_Format, s_Count, s_Posizia As String
                    s_Format = .GetFieldValue_BOM("Формат")
                    s_Count = .GetFieldValue_BOM("Количество")
                    s_Posizia = .GetFieldValue_BOM("Позиция")
                    .CloseBOMItem()
                    combobox_add(ComboBox4)
                    If s_Format IsNot Nothing Then
                        ComboBox4.Items.Add(s_Format)
                        ComboBox4.Text = s_Format
                    End If
                    TextBox6.ReadOnly = False : TextBox6.Text = s_Count : TextBox6.ReadOnly = True
                    TextBox12.ReadOnly = False : TextBox12.Text = s_Posizia : TextBox12.ReadOnly = True
                Catch ex As Exception
                    MsgBox("Заполнении свойств состава" + vbNewLine & ex.Message)
                End Try
            End With
        End If
    End Sub
    Sub combobox_add_Numeric(CB As ComboBox, Num_Count As Integer)
        CB.Items.Clear()
        For Num As Integer = 0 To Num_Count
            CB.Items.Add(Num)
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        read_only(False)
        Button2.Enabled = True
        Button3.Enabled = False
    End Sub
    Sub change_art_doc()
        With Form1.ITS4App
            .OpenArticle(ArtID)
            .SetFieldValue_Articles("Обозначение", a_oboz.Text)
            .SetFieldValue_Articles("Наименование", a_naim.Text)
            .SetFieldValue_Articles("Масса", a_mass.Text)
            '.SetFieldValue_Articles("Материал", TextBox3.Text)
            .CloseArticle()

            If ArtID > 0 Then
                .OpenDocument(DocID)
                .SetFieldValue("Обозначение", d_oboz_TB.Text)
                .SetFieldValue("Наименование", d_naim_TB.Text)
                .SetFieldValue("Формат", d_oboz_TB.Text)
            End If
        End With
    End Sub


    Sub read_only(action As Boolean)
        a_oboz.ReadOnly = action
        a_naim.ReadOnly = action
        a_mass.ReadOnly = action
        a_tolsh.ReadOnly = action

        d_oboz_TB.ReadOnly = action
        d_naim_TB.ReadOnly = action
        d_format_CB.Enabled = Not (action)
        d_Izm_CB.Enabled = Not (action)

        ComboBox4.Enabled = Not (action)
        S_Count_TB.ReadOnly = action '= Me.TextBox6
        S_Positio_TB.ReadOnly = action ' = Me.TextBox12
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With Form1.ITS4App
            .OpenArticle(ArtID)
            .SetFieldValue_Articles("Обозначение", a_oboz.Text)
            .SetFieldValue_Articles("Наименование", a_naim.Text)
            Call .SetFieldValue_Articles("Масса", a_mass.Text)
            Call .SetFieldValue_Articles("Толщина/Диаметр", a_tolsh.Text)
            'Call .SetFieldValue_Articles("Материал", TextBox3.Text)
            .CloseArticle()
            If DocID > 0 Then
                .OpenDocument(DocID)
                .SetFieldValue("Обозначение", d_oboz_TB.Text)
                .SetFieldValue("Наименование", d_naim_TB.Text)
                .SetFieldValue("Формат", d_format_CB.Text)
                .SetDocVersionCode(d_Izm_CB.Text)

                .CloseDocument()
            End If
            If Convert.ToBoolean(get_reesrt_value("ParamSostavVisible", False.ToString)) = True Then
                SetFieldValue_BOM()
            End If
        End With
        Button2.Enabled = False
        Button3.Enabled = True
        read_only(True)
    End Sub
    Sub SetFieldValue_BOM()
        With Form1.ITS4App
            Dim prj_ArtID As Integer
            Dim s1 As String = S_Proj_AID_CB.SelectedItem
            Dim s_result As String
            Try
                If s1.IndexOf(str_ArdID) <> -1 Then
                    s_result = s1.Substring(0, s1.LastIndexOf(separator))
                    s_result = s_result.Substring(s_result.LastIndexOf(" ") + 1)
                End If
                prj_ArtID = .GetArtID_ByDesignation(s_result)
                Dim proj_link_ID As String
                Try
                    .OpenQuery("select * from pc where part_aid = " & ArtID & " and proj_aid = " & prj_ArtID)
                    proj_link_ID = .QueryFieldByName("PRJLINK_ID")
                    .CloseQuery()
                Catch ex As Exception
                    MsgBox("Ошибка при отправке SQL-запроса" + vbNewLine & ex.Message)
                End Try

                If prj_ArtID > 0 Then
                    .OpenBOMItem(proj_link_ID)
                    .SetFieldValue_BOM("Формат", ComboBox4.Text)
                    .SetFieldValue_BOM("Позиция", TextBox12.Text)
                    .SetFieldValue_BOM("Количество", TextBox6.Text)
                End If
            Catch ex As Exception
                MsgBox("Ошибка при изменении свойств состава" + vbNewLine & ex.Message)
            End Try
        End With
    End Sub

    Private Sub Art_property_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.KeyPreview = False
            Me.Close()
        End If
        Me.KeyPreview = True
    End Sub


    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Dim s1 As String = S_Proj_AID_CB.SelectedItem
        If s1 Is Nothing Or s1 = "" Then
            ToolStripButton5.Enabled = False
            Exit Sub
        Else
            ToolStripButton5.Enabled = True
        End If
        'Dim s_result As String
        'If s1.IndexOf(str_ArdID) <> -1 Then
        '    s_result = s1.Substring(0, s1.LastIndexOf(separator))
        '    s_result = s_result.Substring(s_result.LastIndexOf(" ") + 1)
        'End If
        Param_Sostav(GetOboznByComboBox)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Try
            If DocID > 0 Then
                With Form1.ITS4App
                    .OpenDocument(DocID)
                    .View()
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        SetFieldValue_BOM()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try
            If DocID > 0 Then
                With Form1.ITS4App
                    .OpenDocument(DocID)
                    .View()
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Try
            Form1.ITS4App.SetActiveArticle(ArtID)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        With Form1.ITS4App
            Dim oboz As String = GetOboznByComboBox()
            If oboz IsNot Nothing Then
                Dim Pj_ArtID As Integer = .GetArtID_ByDesignation(GetOboznByComboBox)
                .SelectArticleFromTree(ArtID)
            End If
        End With
    End Sub
    Public Function GetOboznByComboBox() As String
        Dim s1 As String = S_Proj_AID_CB.SelectedItem
        Try
            Dim s_result As String
            If s1.IndexOf(str_ArdID) <> -1 Then
                s_result = s1.Substring(0, s1.LastIndexOf(separator))
                s_result = s_result.Substring(s_result.LastIndexOf(" ") + 1)
            End If
            Return s_result
        Catch ex As Exception

        End Try
    End Function
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        With Form1.ITS4App
            .OpenArticle(ArtID)
            .EditParameters_Article()
            .CloseArticle()
        End With
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        With Form1.ITS4App
            .OpenArticle(ArtID)
            .EditParameters_Article()
            .CloseArticle()
        End With
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Try
            With Form1.ITS4App
                Dim thisArtID As Integer = ArtID
                Dim thisDocID As Integer = .GetDocID_ByArtID(thisArtID)
                If thisDocID > 0 Then
                    .OpenDocument(thisDocID)
                    .CheckOut()
                    .EditParameters()
                Else
                    .OpenArticle(thisArtID)
                    .EditParameters_Article()
                    .CloseArticle()
                End If
            End With
        Catch ex As Exception
            MsgBox("Ошибка при попытке открыть карточку документа в S4" + vbNewLine & ex.Message)
        End Try
    End Sub
End Class