Public Class QueryContainer
    Public Event searchQuery()
    Public Event exportJSON()
    Public Event exportCSV()
    Public Event exportXLS()

    Private Sub lblDescript_Click(sender As Object, e As EventArgs) Handles lblDescript.Click

    End Sub

    Private Sub QueryContainer_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        RaiseEvent searchQuery()
    End Sub



    Private Sub btnJson_Click(sender As Object, e As EventArgs) Handles btnJson.Click
        RaiseEvent exportJSON()

    End Sub

    Private Sub btnXls_Click(sender As Object, e As EventArgs) Handles btnXls.Click
        RaiseEvent exportXLS()
    End Sub

    Private Sub btnCsv_Click(sender As Object, e As EventArgs) Handles btnCsv.Click
        RaiseEvent exportCSV()
    End Sub

End Class
