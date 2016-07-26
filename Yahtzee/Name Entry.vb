Public Class frmNameEntry
    Private Sub btnConfirmName_Click(sender As Object, e As EventArgs) Handles btnConfirmName.Click
        If lblFirstName IsNot Nothing Then
            frmYahtzee.lblPlayerName.Text = lblFirstName.Text
            Close()
        End If
    End Sub
End Class