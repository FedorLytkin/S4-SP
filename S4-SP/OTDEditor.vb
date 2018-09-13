Public Class OTDEditor
    Private Sub OTDEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Не регистрирован"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OTD_reg, OTD_Annul As String
        OTD_reg = MaskedTextBox1.Text
        OTD_Annul = MaskedTextBox2.Text
        Dim OTD_cmd As Integer
        Select Case ComboBox1.Text
            Case "Не регистрирован"
                OTD_cmd = 0
                OTD_reg = ""
                OTD_Annul = ""
            Case "Зарегистрирован"
                OTD_cmd = 1
                OTD_Annul = ""
            Case "Аннулирован"
                OTD_cmd = 2
        End Select
        Form1.SetOTDRegnum_SelectDocs(OTD_cmd, OTD_reg, OTD_Annul)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MonthCalendar1.Visible = Not (MonthCalendar1.Visible)
    End Sub


    'Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
    '    TextBox1.Text = MonthCalendar1.SelectionStart
    '    MonthCalendar1.
    'End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.Text
            Case "Зарегистрирован"
                MaskedTextBox1.Enabled = True
                MaskedTextBox1.Text = MonthCalendar1.TodayDate
                MaskedTextBox2.Text = Nothing
                MaskedTextBox2.Enabled = False
                Button3.Enabled = True
                Button4.Enabled = False
            Case "Аннулирован"
                MaskedTextBox1.Text = Nothing
                MaskedTextBox1.Enabled = False
                MaskedTextBox2.Enabled = True
                MaskedTextBox2.Text = MonthCalendar2.TodayDate
                Button3.Enabled = False
                Button4.Enabled = True
            Case "Не регистрирован"
                MaskedTextBox1.Text = Nothing
                MaskedTextBox2.Text = Nothing
                MaskedTextBox1.Enabled = False
                MaskedTextBox2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
            Case Else
                MaskedTextBox1.Enabled = True
                MaskedTextBox2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True

        End Select
        MonthCalendar1.Visible = False
        MonthCalendar2.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MonthCalendar2.Visible = Not (MonthCalendar2.Visible)
        'MonthCalendar2.Select
    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        MaskedTextBox1.Text = MonthCalendar1.SelectionStart
        MonthCalendar1.Visible = Not (MonthCalendar1.Visible)
    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar2.DateSelected
        MaskedTextBox2.Text = MonthCalendar2.SelectionStart
        MonthCalendar2.Visible = Not (MonthCalendar2.Visible)
    End Sub
    Public Sub ProcessAdd(ProcessBarString As String)
        ListBox1.Items.Add(ProcessBarString)
    End Sub

    Private Sub ОчиститьЗаписьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОчиститьЗаписьToolStripMenuItem.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub ОткрытьКарточкуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьКарточкуToolStripMenuItem.Click
        Try
            Dim s1 As String = ListBox1.SelectedItem
            If s1.IndexOf("DocID:") < 0 Then Exit Sub

            Dim result As String = s1.Substring(0, s1.IndexOf(" - "))
            result = result.Substring(s1.IndexOf(": ") + 2).Trim
            With Form1.ITS4App
                .OpenDocument(result)
                .EditParameters()
                .CloseDocument()
            End With
        Catch ex As Exception

        End Try
        'Checking_SOSTAV_Form.Edit_Parameters_Article_Proc()
        'Edit_Parameters_Article_Proc()
        'Checking_SOSTAV_Form.open_IProperty()
    End Sub

    Public Function ART_or_DOC_ID() As Integer
        Try
            Dim str_ArdID As String = "DocID: "
            Dim separator As String = " - "

            Dim s1 As String = ListBox1.SelectedItem

            Dim s_result As String
            If s1.IndexOf(str_ArdID) <> -1 Then
                s_result = s1.Substring(0, s1.LastIndexOf(separator))
                s_result = s_result.Substring(s_result.LastIndexOf(" ") + 1)
            End If
            Return s_result
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Sub Edit_Parameters_Article_Proc()
        With Form1.ITS4App
            .OpenArticle(ART_or_DOC_ID())
            .EditParameters_Article()
            .CloseArticle()
        End With
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Process.Start("http://www.evernote.com/l/AjJ2IvTGcmJLlLDt2Q2MP6tP6mS9RSSPpkM/")
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)
        Dim OTD_reg, OTD_Annul As String
        OTD_reg = MaskedTextBox1.Text
        OTD_Annul = MaskedTextBox2.Text
        Dim OTD_cmd As Integer
        Select Case ComboBox1.Text
            Case "Не регистрирован"
                OTD_cmd = 0
                OTD_reg = ""
                OTD_Annul = ""
            Case "Зарегистрирован"
                OTD_cmd = 1
                OTD_Annul = ""
            Case "Аннулирован"
                OTD_cmd = 2
        End Select
        Form1.SetOTDRegnum_SelectDocs(OTD_cmd, OTD_reg, OTD_Annul)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If ListBox1.Items.Count = 0 Then
            ОчиститьЗаписьToolStripMenuItem.Enabled = False
            ОткрытьКарточкуToolStripMenuItem.Enabled = False
        Else
            ОчиститьЗаписьToolStripMenuItem.Enabled = True
            ОткрытьКарточкуToolStripMenuItem.Enabled = True
        End If
    End Sub
End Class