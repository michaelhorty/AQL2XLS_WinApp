Public Class MainUI

    Private Sub MainUI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' main entry point
        Call multiControls()

    End Sub

    Private Sub multiControls()
        If CheckBox1.Checked = True Then
            QueryContainer1.btnGo.Visible = True
            QueryContainer2.btnGo.Visible = True
            QueryContainer2.TextBox1.Visible = True
            QueryContainer2.Visible = True
            QueryContainer2.Label1.Visible = True
            QueryContainer2.btnXls.Visible = True
            QueryContainer2.btnCsv.Visible = True
            QueryContainer2.btnJson.Visible = True
            QueryContainer1.btnXls.Visible = True
            QueryContainer1.btnCsv.Visible = True
            QueryContainer1.btnJson.Visible = True

        Else
            QueryContainer2.btnGo.Visible = True
            QueryContainer1.btnGo.Visible = False
            QueryContainer2.TextBox1.Visible = False
            QueryContainer2.Label1.Visible = False
            QueryContainer2.Visible = True
            QueryContainer2.btnXls.Visible = True
            QueryContainer2.btnCsv.Visible = True
            QueryContainer2.btnJson.Visible = True
            QueryContainer1.btnXls.Visible = False
            QueryContainer1.btnCsv.Visible = False
            QueryContainer1.btnJson.Visible = False


        End If
    End Sub

    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        Call multiControls()

    End Sub
End Class
