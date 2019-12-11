Public Class newOVActl
    Public Event createOVA(ipAddy$, netMask$, gateWay$, ntpServer$, dnsServer$, proxy$, ovaID$)
    Public Event userCancelled()

    Private Sub newOVActl_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
    Public Sub setupCTL(C As Collection)
        cboOva.Items.Clear()

        For Each OVA In C
            cboOva.Items.Add(OVA)
        Next

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        RaiseEvent userCancelled()
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        RaiseEvent createOVA(txtIP.Text, txtMask.Text, txtGW.Text, txtNTP.Text, txtDNS.Text, txtProxy.Text, cboOva.Text)
    End Sub

    Private Sub txtIP_TextChanged(sender As Object, e As EventArgs) Handles txtIP.TextChanged

    End Sub


End Class
