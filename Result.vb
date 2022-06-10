Public Class Result

    Private Sub Result_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        KeyPreview = True
    End Sub

    Private Sub Btn_Create_Click(sender As Object, e As EventArgs) Handles Btn_Create.Click
        Clipboard.SetText(TextBox1.Text)
    End Sub
End Class