Public Class frmLogin
    'Public Event loginCreds(frmUserN$, frmToken$)
    Public loginUser$
    Public loginToken$
    Private loginS As Collection
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With MainUI.authInfo

            .fqdN = "https://" + ComboBox1.Text + ".armis.com"
            .secretKey = TextBox2.Text

        End With

        Me.Close()
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Dim L As Collection
        loginS = getLogins()

        ComboBox1.Items.Clear()

        Dim l$

        For Each logIN In loginS
            l$ = Mid(logIN, 1, InStr(logIN, "|") - 1)
            l = Replace(l, "https://", "")
            l = Replace(l, ".armis.com", "")

            ComboBox1.Items.Add(l)
        Next

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim tenanT$ = ComboBox1.Text

        If Len(tenanT) = 0 Then Exit Sub

        Dim S3 As New Simple3Des("7w6e87twryut24876wuyeg")

        For Each L In loginS
            If LCase("https://" + tenanT + ".armis.com") = LCase(Mid(L, 1, InStr(L, "|") - 1)) Then
                TextBox2.Text = S3.Decode(Mid(L, InStr(L, "|") + 1))
            End If
        Next
    End Sub
End Class