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
                                tmp_color = GetStrToColor(get_reesrt_value("DocumColor_Sostav", GetColorToStr(.DocumColor_Sostav)))
                            Case "Сборочные единицы"
                                tmp_color = GetStrToColor(get_reesrt_value("AssemblyColor_Sostav", GetColorToStr(.AssemblyColor_Sostav)))
                            Case "Детали"
                                tmp_color = GetStrToColor(get_reesrt_value("PartColor_Sostav", GetColorToStr(.PartColor_Sostav)))
                                'tmp_color = .PartColor_Sostav
                            Case "Стандартные изделия"
                                tmp_color = GetStrToColor(get_reesrt_value("StandartColor_Sostav", GetColorToStr(.StandartColor_Sostav)))
                                'tmp_color = .StandartColor_Sostav
                            Case "Прочие изделия"
                                tmp_color = GetStrToColor(get_reesrt_value("Pro4eeColor_Sostav", GetColorToStr(.Pro4eeColor_Sostav)))
                                'tmp_color = .Pro4eeColor_Sostav
                            Case "Материалы"
                                tmp_color = GetStrToColor(get_reesrt_value("MaterialColor_Sostav", GetColorToStr(.MaterialColor_Sostav)))
                                'tmp_color = .MaterialColor_Sostav
                            Case "Комплекты"
                                tmp_color = GetStrToColor(get_reesrt_value("KomplekTColor_Sostav", GetColorToStr(.KomplekTColor_Sostav)))
                                'tmp_color = .KomplekTColor_Sostav
                            Case "Технологич.связь"
                                tmp_color = GetStrToColor(get_reesrt_value("TechContex_Sostav", GetColorToStr(.TechContex_Sostav)))
                                'tmp_color = .TechContex_Sostav
                            Case "Материал вспомогательный"
                                tmp_color = GetStrToColor(get_reesrt_value("VspomMaterialColor_Purchated", GetColorToStr(.VspomMaterialColor_Purchated)))
                                'tmp_color = .VspomMaterialColor_Purchated
                            Case "Сортамент изделия"
                                tmp_color = GetStrToColor(get_reesrt_value("SortamentColor_Purchated", GetColorToStr(.SortamentColor_Purchated)))
                                'tmp_color = .SortamentColor_Purchated
                            Case "Материал изделия(Констр)"
                                tmp_color = GetStrToColor(get_reesrt_value("SortamentColorKonstructor_Purchated", GetColorToStr(.SortamentColorKonstructor_Purchated)))
                                'tmp_color = .SortamentColorKonstructor_Purchated
                            Case "Стандартные изделия"
                                tmp_color = GetStrToColor(get_reesrt_value("StandartColor_Purchated", GetColorToStr(.StandartColor_Purchated)))
                                'tmp_color = .StandartColor_Purchated
                            Case "Прочие изделия"
                                tmp_color = GetStrToColor(get_reesrt_value("Pro4eeColor_Purchated", GetColorToStr(.Pro4eeColor_Purchated)))
                                'tmp_color = .Pro4eeColor_Purchated
                            Case "Материал"
                                tmp_color = GetStrToColor(get_reesrt_value("MaterialColor_Purchated", GetColorToStr(.MaterialColor_Purchated)))
                                'tmp_color = .MaterialColor_Purchated
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
        ColorDialog1.FullOpen = 1
        ColorDialog1.Color = DownButton.BackColor
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
        Dim RegParName As String
        For i As Integer = 0 To Me.TabControl1.Controls.Count - 1
            For Each tmp_butt In Me.TabControl1.Controls.Item(i).Controls
                If (tmp_butt.GetType.ToString = "System.Windows.Forms.Button") Then
                    tmp_color = tmp_butt.BackColor
                    With VSNRM2
                        Select Case tmp_butt.Text
                            Case "Документация"
                                .DocumColor_Sostav = tmp_color
                                RegParName = "DocumColor_Sostav"
                            Case "Сборочные единицы"
                                .AssemblyColor_Sostav = tmp_color
                                RegParName = "AssemblyColor_Sostav"
                            Case "Детали"
                                .PartColor_Sostav = tmp_color
                                RegParName = "PartColor_Sostav"
                            Case "Стандартные изделия"
                                .StandartColor_Sostav = tmp_color
                                RegParName = "StandartColor_Sostav"
                            Case "Прочие изделия"
                                .Pro4eeColor_Sostav = tmp_color
                                RegParName = "Pro4eeColor_Sostav"
                            Case "Материалы"
                                .MaterialColor_Sostav = tmp_color
                                RegParName = "MaterialColor_Sostav"
                            Case "Комплекты"
                                .KomplekTColor_Sostav = tmp_color
                                RegParName = "KomplekTColor_Sostav"
                            Case "Технологич.связь"
                                .TechContex_Sostav = tmp_color
                                RegParName = "TechContex_Sostav"
                            Case "Материал вспомогательный"
                                .VspomMaterialColor_Purchated = tmp_color
                                RegParName = "VspomMaterialColor_Purchated"
                            Case "Сортамент изделия"
                                .SortamentColor_Purchated = tmp_color
                                RegParName = "SortamentColor_Purchated"
                            Case "Материал изделия(Констр)"
                                .SortamentColorKonstructor_Purchated = tmp_color
                                RegParName = "SortamentColorKonstructor_Purchated"
                            Case "Стандартные изделия"
                                .StandartColor_Purchated = tmp_color
                                RegParName = "StandartColor_Purchated"
                            Case "Прочие изделия"
                                .Pro4eeColor_Purchated = tmp_color
                                RegParName = "Pro4eeColor_Purchated"
                            Case "Материал"
                                .MaterialColor_Purchated = tmp_color
                                RegParName = "MaterialColor_Purchated"
                            Case Else
                                tmp_color = Color.Transparent
                        End Select
                        set_reesrt_value(RegParName, GetColorToStr(tmp_color))

                    End With
                End If
            Next
        Next
        MsgBox("Изменения сохранены!")
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs)
        ColorDialog1.FullOpen = 1
        ColorDialog1.Color = VSNRM2.AssemblyColor_Sostav
        Dim color_str As String = ColorDialog1.Color.ToString
        ColorDialog1.Color = Color.FromName(color_str)
        If ColorDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            DownButton.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs)
        Dim a, r, g, b As Byte
        a = 255
        r = 216
        g = 228
        b = 188
        Button17.BackColor = Color.FromArgb(a, r, g, b)
    End Sub
    Function GetColorToStr(tmp_color As Color) As String
        Dim color_str As String = tmp_color.ToString
        Dim argb() As String
        Dim a, r, g, b As String
        Try
            color_str = color_str.Replace("Color", "").Trim(" ").Trim("[").Trim("]")
            If color_str.IndexOf(",") > 0 Then
                argb = color_str.Split(",")
                a = argb(0).Replace("A=", "").Trim(",")
                r = argb(1).Replace("R=", "").Trim(",")
                g = argb(2).Replace("G=", "").Trim(",")
                b = argb(3).Replace("B=", "").Trim(",")
                Return a & "," & r & "," & g & "," & b
            Else
                Return color_str
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetStrToColor(clr_str As String) As Color
        Dim argb() As String
        Dim a, r, g, b As Byte

        Try
            If clr_str.IndexOf(",") > 0 Then
                argb = clr_str.Split(",")
                a = CByte(argb(0).Replace("A=", "").Trim(","))
                r = CByte(argb(1).Replace("R=", "").Trim(","))
                g = CByte(argb(2).Replace("G=", "").Trim(","))
                b = CByte(argb(3).Replace("B=", "").Trim(","))
                Return Color.FromArgb(a, r, g, b)
            Else
                Return Color.FromName(clr_str)
            End If
        Catch ex As Exception
        End Try
    End Function

    Private Sub Button17_Click_1(sender As Object, e As EventArgs) Handles Button17.Click
        Dim tmp_butt As Control
        Dim tmp_color As Color
        Dim RegParName As String
        For i As Integer = 0 To Me.TabControl1.Controls.Count - 1
            For Each tmp_butt In Me.TabControl1.Controls.Item(i).Controls
                If (tmp_butt.GetType.ToString = "System.Windows.Forms.Button") Then
                    With VSNRM2
                        Select Case tmp_butt.Text
                            Case "Документация"
                                tmp_color = .DocumColor_Sostav
                                RegParName = "DocumColor_Sostav"
                            Case "Сборочные единицы"
                                tmp_color = .AssemblyColor_Sostav
                                RegParName = "AssemblyColor_Sostav"
                            Case "Детали"
                                tmp_color = .PartColor_Sostav
                                RegParName = "PartColor_Sostav"
                            Case "Стандартные изделия"
                                tmp_color = .StandartColor_Sostav
                                RegParName = "StandartColor_Sostav"
                            Case "Прочие изделия"
                                tmp_color = .Pro4eeColor_Sostav
                                RegParName = "Pro4eeColor_Sostav"
                            Case "Материалы"
                                tmp_color = .MaterialColor_Sostav
                                RegParName = "MaterialColor_Sostav"
                            Case "Комплекты"
                                tmp_color = .KomplekTColor_Sostav
                                RegParName = "KomplekTColor_Sostav"
                            Case "Технологич.связь"
                                tmp_color = .TechContex_Sostav
                                RegParName = "TechContex_Sostav"
                            Case "Материал вспомогательный"
                                tmp_color = .VspomMaterialColor_Purchated
                                RegParName = "VspomMaterialColor_Purchated"
                            Case "Сортамент изделия"
                                tmp_color = .SortamentColor_Purchated
                                RegParName = "SortamentColor_Purchated"
                            Case "Материал изделия(Констр)"
                                tmp_color = .SortamentColorKonstructor_Purchated
                                RegParName = "SortamentColorKonstructor_Purchated"
                            Case "Стандартные изделия"
                                tmp_color = .StandartColor_Purchated
                                RegParName = "StandartColor_Purchated"
                            Case "Прочие изделия"
                                tmp_color = .Pro4eeColor_Purchated
                                RegParName = "Pro4eeColor_Purchated"
                            Case "Материал"
                                tmp_color = .MaterialColor_Purchated
                                RegParName = "MaterialColor_Purchated"
                            Case Else
                                tmp_color = Color.Transparent
                        End Select
                        set_reesrt_value(RegParName, GetColorToStr(tmp_color))
                        tmp_butt.BackColor = tmp_color
                    End With
                End If
            Next
        Next
        MsgBox("Изменения по умолчанию сохранены!")
    End Sub
End Class