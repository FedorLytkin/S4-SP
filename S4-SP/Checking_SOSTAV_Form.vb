Public Class Checking_SOSTAV_Form
    Public ArtID As Integer
    Public DocID As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub СправкаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СправкаToolStripMenuItem.Click
        Process.Start("http://www.evernote.com/l/AjJB4VajI-hLr6r9Zbxjpkd5lY9hHpZN0C4/")
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Form1.check_article_SOSTAV()
    End Sub

    Private Sub ПоказатьВS4ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоказатьВS4ToolStripMenuItem.Click
        SelectInS4()
    End Sub
    Public Sub SelectInS4()
        Form1.ITS4App.SetActiveArticle(ART_or_DOC_ID)
    End Sub
    Public Function ART_or_DOC_ID() As Integer
        Try
            Dim str_ArdID As String = "ArtID: "
            Dim separator As String = " - "

            Dim s1 As String = ListBox1.SelectedItem

            Dim s_result As String
            If s1.IndexOf(str_ArdID) <> -1 Then
                s_result = s1.Substring(0, s1.LastIndexOf(separator))
                s_result = s_result.Substring(s_result.LastIndexOf(" ") + 1)
            End If
            Return s_result
        Catch ex As Exception

        End Try
    End Function

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)
        ART_or_DOC_ID()
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        buttn_enabled()
        'Dim str_ArdID As String = "ArtID: "
        'Dim separator As String = " - "
        'Dim s1 As String = ListBox1.SelectedItem
        'Try
        '    If s1.IndexOf(str_ArdID) > -1 Then
        '        ОткрытьКарточкуToolStripMenuItem.Enabled = True
        '        ПоказатьВS4ToolStripMenuItem.Enabled = True
        '    Else
        '        ОткрытьКарточкуToolStripMenuItem.Enabled = False
        '        ПоказатьВS4ToolStripMenuItem.Enabled = False
        '    End If
        'Catch ex As Exception

        'End Try
    End Sub
    Sub buttn_enabled()
        Dim str_ArdID As String = "ArtID: "
        Dim separator As String = " - "
        Dim s1 As String = ListBox1.SelectedItem
        Try
            If s1.IndexOf(str_ArdID) > -1 Then
                ОткрытьКарточкуToolStripMenuItem.Enabled = True
                ПоказатьВS4ToolStripMenuItem.Enabled = True
                ToolStripButton2.Enabled = True
                ToolStripButton3.Enabled = True
                ToolStripSplitButton2.Enabled = True
                ToolStripMenuItem2.Enabled = True
            Else
                ОткрытьКарточкуToolStripMenuItem.Enabled = False
                ПоказатьВS4ToolStripMenuItem.Enabled = False
                ToolStripButton2.Enabled = False
                ToolStripButton3.Enabled = False
                ToolStripSplitButton2.Enabled = False
                ToolStripMenuItem2.Enabled = False
            End If
        Catch ex As Exception
            ОткрытьКарточкуToolStripMenuItem.Enabled = False
            ПоказатьВS4ToolStripMenuItem.Enabled = False
            ToolStripButton2.Enabled = False
            ToolStripButton3.Enabled = False
            ToolStripSplitButton2.Enabled = False
            ToolStripMenuItem2.Enabled = False
        End Try
        If Form1.Carto4kaS4SPEnable = False Then
            ToolStripButton2.Enabled = False
        Else
            ToolStripButton2.Enabled = True
        End If
    End Sub

    Private Sub ОткрытьКарточкуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьКарточкуToolStripMenuItem.Click
        open_IProperty()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        open_IProperty()
    End Sub

    Sub open_IProperty()
        Art_property.ArtID = ART_or_DOC_ID()
        Art_property.ShowDialog()
    End Sub

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        buttn_enabled()
    End Sub

    Private Sub ListBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseClick
        buttn_enabled()
    End Sub

    Private Sub УдалитьЗаписчьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьЗаписчьToolStripMenuItem.Click
        ListBox1.Items.Remove(ListBox1.SelectedItem)
    End Sub

    Private Sub ОчиститьПолеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОчиститьПолеToolStripMenuItem.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Form1.check_article_SOSTAV()
    End Sub

    Private Sub ToolStripButton3_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        SelectInS4()
    End Sub

    Private Sub ОткрытьКарточкуДокументаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьКарточкуДокументаToolStripMenuItem.Click
        Edit_Parameters_Article_Proc()
    End Sub
    Sub Edit_Parameters_Article_Proc()
        With Form1.ITS4App
            .OpenArticle(ART_or_DOC_ID())
            .EditParameters_Article()
            .CloseArticle()
        End With
    End Sub

    Private Sub ПросмотрКарточкиДокументаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПросмотрКарточкиДокументаToolStripMenuItem.Click
        Edit_Parameters_Article_Proc()
    End Sub

    Private Sub РедактированиеКарточкиДокументаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles РедактированиеКарточкиДокументаToolStripMenuItem.Click
        EditParametersAndCheckOutDocument()
    End Sub
    Sub EditParametersAndCheckOutDocument()
        With Form1.ITS4App
            Dim thisArtID As Integer = ART_or_DOC_ID()
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
    End Sub

    Private Sub РедактированиеКарточкиДокументаToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles РедактированиеКарточкиДокументаToolStripMenuItem1.Click
        EditParametersAndCheckOutDocument()
    End Sub

    Private Sub Checking_SOSTAV_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form1.Carto4kaS4SPEnable = False Then
            ToolStripButton2.Enabled = False
        Else
            ToolStripButton2.Enabled = True
        End If
    End Sub
End Class