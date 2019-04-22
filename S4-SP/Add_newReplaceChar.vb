Public Class Add_newReplaceChar
    Private Sub Cancel_Bt_Click(sender As Object, e As EventArgs) Handles Cancel_Bt.Click
        OK_Bt.Text = ""
        Cancel_Bt.Text = ""
        Close()
    End Sub

    Private Sub Add_newReplaceChar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AcceptButton = OK_Bt
        Me.CancelButton = Cancel_Bt
    End Sub

End Class