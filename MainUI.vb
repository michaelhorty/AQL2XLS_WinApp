Public Class MainUI
    Public loggingEnabled As Boolean
    Public guiActive As Boolean
    Delegate Sub StringArgReturningVoidDelegate([text] As String)
    Public authInfo As apiAuthInfo
    Private Sub MainUI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' main entry point
        Loggingenabled = True
        guiActive = True
        Call multiControls()
        addLOG("GUI Activated")

        authInfo = New apiAuthInfo
        With authInfo
            .fqdN = "https://" + TextBox1.Text + ".armis.com"
            .secretKey = TextBox2.Text
        End With

        Dim newClient As New ARMclient(authInfo)

        If newClient.gotToken = False Then
            addLOG("Error getting token:" + newClient.lastError)
        Else
            addLOG("Token: " + authInfo.tokeN)
            addLOG("Expires: UTC " + authInfo.tokenExpires.ToString("hh:mm:ss"))
        End If

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
    Public Sub addLOG(ByVal a$, Optional ByVal suppressDT As Boolean = False, Optional ByVal forceLog As Boolean = False)
        On Error GoTo errorcatch
        'logginGenabled = True


forFileOnly:
        If suppressDT = False Then a = Now.ToString("MM/dd hh:mm:ss") + ": " + a

        Dim fileN$ = "armxls_log.txt"

        If loggingEnabled = False And forceLog = False Then GoTo writelineOnly

        Dim FF As Integer = FreeFile()

        If Dir(fileN) = "" Then
            FileOpen(FF, fileN, OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        Else
            FileOpen(FF, fileN, OpenMode.Append, OpenAccess.Write, OpenShare.Shared)
        End If


        Print(FF, a + vbCrLf)
        FileClose(FF)

writelineOnly:
        If guiActive = True Then logTEXT(a + vbCrLf)
errorcatch:
    End Sub

    Public Sub logTEXT(ByVal a$)
        If Me.logTB.InvokeRequired Then
            'this delegate nonsense is to ensure background thread request is accepted by the logTB control
            Dim d As New StringArgReturningVoidDelegate(AddressOf logTEXT)
            Me.Invoke(d, New Object() {a})
        Else
            logTB.AppendText(a)
            logTB.SelectionStart = logTB.Text.Length
            logTB.ScrollToCaret()
            logTB.Refresh()
        End If
    End Sub

    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        Call multiControls()

    End Sub

    Private Sub QueryContainer1_searchQuery() Handles QueryContainer1.searchQuery, QueryContainer2.searchQuery
        'ently, the following types of searches are supported:
        '1. in:Devices
        '2. in:alerts
        '3. In:activity
        Dim qrY$

        qrY = QueryContainer1.TextBox1.Text

        Call submitQuery(qrY)

    End Sub

    Private Sub submitQuery(qry)
        If authInfo.tokeN = "" Then
            MsgBox("Not authenticated.. Check your Secret Key!", vbOKOnly, "Not Connected")
            Exit Sub
        End If

        Dim AClient As New ARMclient(authInfo)
        AClient.searchAPI(qry)


    End Sub

End Class
