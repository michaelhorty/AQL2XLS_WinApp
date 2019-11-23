Public Class QueryContainer
    Public Event searchQuery()


    Private Sub lblDescript_Click(sender As Object, e As EventArgs) Handles lblDescript.Click

    End Sub

    Private Sub QueryContainer_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        RaiseEvent searchQuery()
    End Sub
End Class
