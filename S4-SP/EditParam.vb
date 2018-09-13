Public Class EditParam
    Public fldName, fldValue As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With Form1.ITS4App
            Dim AID, DID As Integer
            Dim SelItems As S4.TS4SelectedItems = .GetSelectedItems
            While SelItems.NextSelected <> 0
                DID = SelItems.ActiveDocID
                If DID > 0 Then
                    AID = .GetArtID_ByDocID(DID)
                Else
                    AID = SelItems.ActiveArtID
                End If
                If AID > 0 Then
                    .SetFieldValue(ComboBox1.Text, TextBox1.Text)
                End If
                SelItems.InvertCurrent()
            End While
        End With
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        With Form1.ITS4App
            fldName = ComboBox1.Text
            TextBox2.Text = .GetFieldValue(fldName)
        End With
    End Sub

    Private Sub EditParam_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = ""
        ComboBox1.Items.Clear()
        With Form1.ITS4App
            Dim DID, fldCount As Integer
            Dim SelItems As S4.TS4SelectedItems = .GetSelectedItems
            DID = SelItems.ActiveDocID
            .OpenDocument(DID)
            fldCount = .GetFieldCount
            For i As Integer = 0 To fldCount - 1
                fldName = .GetFieldName(i)
                ComboBox1.Items.Add(fldName)
            Next
            'While SelItems.NextSelected <> 0
            '    DID = SelItems.ActiveDocID
            '    If DID > 0 Then
            '        AID = .GetArtID_ByDocID(DID)
            '    Else
            '        AID = SelItems.ActiveArtID
            '    End If
            '    If AID > 0 Then

            '    End If
            '    SelItems.InvertCurrent()
            'End While
        End With
        ComboBox1.Text = ComboBox1.Items.Item(0)
    End Sub
End Class