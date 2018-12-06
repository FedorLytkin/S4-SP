Public Class ColorOptions
    Private Sub ColorOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tmp_butt As Control
        Dim tmp_color As Color
        For i As Integer = 0 To Me.TabControl1.Controls.Count - 1
            For Each tmp_butt In Me.TabControl1.Controls.Item(i).Controls
                If (tmp_butt.GetType.ToString = "System.Windows.Forms.Button") Then

                    With VSNRM2
                        Select Case tmp_butt.Text
                            Case "Документация"
                                tmp_color = .DocumColor_Sostav
                            Case "Сборочные единицы"
                                tmp_color = .AssemblyColor_Sostav
                            Case "Детали"
                                tmp_color = .PartColor_Sostav
                            Case "Стандартные изделия"
                                tmp_color = .StandartColor_Sostav
                            Case "Прочие изделия"
                                tmp_color = .Pro4eeColor_Sostav
                            Case "Материалы"
                                tmp_color = .MaterialColor_Sostav
                            Case "Комплекты"
                                tmp_color = .KomplekTColor_Sostav
                            Case "Технологич.связь"
                                tmp_color = .TechContex_Sostav
                            Case "Материал вспомогательный"
                                tmp_color = .VspomMaterialColor_Purchated
                            Case "Сортамент изделия"
                                tmp_color = .SortamentColor_Purchated
                            Case "Материал изделия(Констр)"
                                tmp_color = .SortamentColorKonstructor_Purchated
                            Case "Стандартные изделия"
                                tmp_color = .StandartColor_Purchated
                            Case "Прочие изделия"
                                tmp_color = .Pro4eeColor_Purchated
                            Case "Материал"
                                tmp_color = .MaterialColor_Purchated
                            Case Else
                                tmp_color = Color.Transparent
                        End Select
                    End With
                    tmp_butt.BackColor = tmp_color
                End If
            Next
        Next
    End Sub
    Dim DownButton As Button
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DownButton = Button9
        colorDialogLoad()
    End Sub
    Sub colorDialogLoad()
        If ColorDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            DownButton.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        DownButton = Button10
        colorDialogLoad()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        DownButton = Button11
        colorDialogLoad()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        DownButton = Button12
        colorDialogLoad()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        DownButton = Button13
        colorDialogLoad()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        DownButton = Button14
        colorDialogLoad()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DownButton = Button1
        colorDialogLoad()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DownButton = Button2
        colorDialogLoad()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DownButton = Button3
        colorDialogLoad()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DownButton = Button4
        colorDialogLoad()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DownButton = Button5
        colorDialogLoad()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DownButton = Button6
        colorDialogLoad()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DownButton = Button7
        colorDialogLoad()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        DownButton = Button8
        colorDialogLoad()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Me.Close()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim tmp_butt As Control
        Dim tmp_color As Color
        For i As Integer = 0 To Me.TabControl1.Controls.Count - 1
            For Each tmp_butt In Me.TabControl1.Controls.Item(i).Controls
                If (tmp_butt.GetType.ToString = "System.Windows.Forms.Button") Then
                    tmp_color = tmp_butt.BackColor
                    With VSNRM2
                        Select Case tmp_butt.Text
                            Case "Документация"
                                .DocumColor_Sostav = tmp_color
                            Case "Сборочные единицы"
                                .AssemblyColor_Sostav = tmp_color
                            Case "Детали"
                                .PartColor_Sostav = tmp_color
                            Case "Стандартные изделия"
                                .StandartColor_Sostav = tmp_color
                            Case "Прочие изделия"
                                .Pro4eeColor_Sostav = tmp_color
                            Case "Материалы"
                                .MaterialColor_Sostav = tmp_color
                            Case "Комплекты"
                                .KomplekTColor_Sostav = tmp_color
                            Case "Технологич.связь"
                                .TechContex_Sostav = tmp_color
                            Case "Материал вспомогательный"
                                .VspomMaterialColor_Purchated = tmp_color
                            Case "Сортамент изделия"
                                .SortamentColor_Purchated = tmp_color
                            Case "Материал изделия(Констр)"
                                .SortamentColorKonstructor_Purchated = tmp_color
                            Case "Стандартные изделия"
                                .StandartColor_Purchated = tmp_color
                            Case "Прочие изделия"
                                .Pro4eeColor_Purchated = tmp_color
                            Case "Материал"
                                .MaterialColor_Purchated = tmp_color
                            Case Else
                                tmp_color = Color.Transparent
                        End Select
                    End With
                End If
            Next
        Next
        MsgBox("Изменения сохранены!")
    End Sub
End Class