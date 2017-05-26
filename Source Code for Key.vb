Public Class Key
    Public Property TxtKey As Object

    Private Sub ButCopy_Click(sender As Object, e As EventArgs) Handles ButCopy.Click
        Clipboard.SetText(TxtKey.Text)
    End Sub
End Class
