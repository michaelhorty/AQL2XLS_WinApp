Public Class frmRptDevice
    Private Sub frmRptDevice_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Public Sub popListBox(ByRef DevCol As Collection, ByRef TagCol As Collection)
        'devices default to CHECKED for first 16
        Dim N As Integer = 1

        cbLB.Items.Clear()

        For Each D In DevCol
            cbLB.Items.Add(D, True)
        Next

        cbLB.Items.Add("Tag List", True)

        For Each T In TagCol
            cbLB.Items.Add("TAG:" + T, False)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub btnUP_Click(sender As Object, e As EventArgs) Handles btnUP.Click
        Dim nDX = cbLB.SelectedIndex

        If nDX = -1 Then Exit Sub

        If nDX > 0 Then
            Dim seL$ = cbLB.Items(nDX).ToString
            Dim swP$ = cbLB.Items(nDX - 1).ToString

            cbLB.Items(nDX) = swP
            cbLB.Items(nDX - 1) = seL

            cbLB.SelectedIndex -= 1
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim nDX = cbLB.SelectedIndex

        If nDX = -1 Then Exit Sub

        If nDX < cbLB.Items.Count - 1 Then
            Dim seL$ = cbLB.Items(nDX).ToString
            Dim swP$ = cbLB.Items(nDX + 1).ToString

            cbLB.Items(nDX) = swP
            cbLB.Items(nDX + 1) = seL

            cbLB.SelectedIndex += 1
        End If
    End Sub
End Class