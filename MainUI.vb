Imports System.Threading

Public Class MainUI
    Private loggingEnabled As Boolean
    Private guiActive As Boolean
    Delegate Sub StringArgReturningVoidDelegate([text] As String)
    Private authInfo As apiAuthInfo

    Private jsoN1$
    Private jsoN2$
    Private deviceS1 As List(Of deviceData)

    Private objType1 As queryType
    Private objType2 As queryType


    Private WithEvents AClient As ARMclient

    Private processTasks1 As Long 'next query to be written to deviceS1
    Private processTasks2 As Long
    '    Private processTasksDone1 As Collection
    '    Private processTasksDone2 As Collection

    Private Sub MainUI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' main entry point
        loggingEnabled = True
        guiActive = True
        Call multiControls()
        addLOG("GUI Activated")

        authInfo = New apiAuthInfo
        With authInfo
            .fqdN = "https://" + TextBox1.Text + ".armis.com"
            .secretKey = TextBox2.Text
        End With

        Dim newClient As New ARMclient(authInfo)

        If Len(authInfo.tokeN) = 0 Then
            addLOG("Error getting token:" + newClient.lastError)
        Else
            addLOG("Token: " + authInfo.tokeN)
            addLOG("Expires: UTC " + authInfo.tokenExpires.ToString("hh:mm:ss"))
        End If

        newClient = Nothing
    End Sub

    Private Sub multiControls()

        QueryContainer2.Visible = CheckBox1.Checked


        Exit Sub

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

    Private Sub QueryContainer1_searchQuery() Handles QueryContainer1.searchQuery ', QueryContainer2.searchQuery
        'ently, the following types of searches are supported:
        '1. in:Devices
        '2. in:alerts
        '3. In:activity
        Dim qrY$

        qrY = QueryContainer1.TextBox1.Text

        'QueryContainer1.btnGo.Visible = False

        Call submitQuery(qrY, 1)

    End Sub


    Private Sub submitQuery(qry$, Optional ByVal queryBox1or2 As Integer = 1)
        If authInfo.tokeN = "" Then
            MsgBox("Not authenticated.. Check your Secret Key!", vbOKOnly, "Not Connected")
            Exit Sub
        End If

        AClient = New ARMclient(authInfo)

        objType1 = queryType.Undefined
        objType2 = queryType.Undefined

        If LCase(Mid(qry, 1, 10)) = "in:devices" Then objType1 = queryType.Devices
        If LCase(Mid(qry, 1, 13)) = "in:activities" Then objType1 = queryType.Activities
        If LCase(Mid(qry, 1, 9)) = "in:alerts" Then objType1 = queryType.Alerts

        If objType1 = queryType.Undefined Then
            MsgBox("SEARCH queries support Devices, Activities and Alerts. Please enter a valid query.", vbOKOnly, "Invalid Query")
        End If

        If queryBox1or2 = 1 Then
            Call enableButtons(1, True)
            QueryContainer1.btnGo.Visible = False
        Else
            Call enableButtons(2, True)
            QueryContainer2.btnGo.Visible = False
        End If


        Dim jsonText$ = ""

        Dim nextResultStart As Integer = 0
        Dim numResultsPerCall As Integer = 100
        Dim tlNum As Long = 0
        deviceS1 = New List(Of deviceData)


        Dim mThreading As threadingArgs
        mThreading = New threadingArgs(authInfo)
        mThreading.qrY = qry

        jsonText = AClient.searchAPI(mThreading)

        If queryBox1or2 = 1 Then jsoN1 = jsonText Else jsoN2 = jsonText

        Dim respData As New deviceResponseData
        respData = AClient.deserializeResponseData(jsoN1)

        tlNum = respData.total
        If tlNum <= 100 Then GoTo allDone

        addLOG("Queuing " + Math.Round((tlNum - 100) / 100, 0).ToString + " API threads for " + tlNum.ToString + " devices")

        Dim w As WaitCallback = New WaitCallback(AddressOf AClient.searchAPI)


        'depending on reliability
        'may need to identify requests that werent answered
        'if queuing changed from 200 to 100ms, UI seems to not return responses

        nextResultStart = respData.count

        ThreadPool.SetMaxThreads(100, 100)

        Dim numMS As Long = 0
        Dim startReq As DateTime = Now

        Do Until nextResultStart >= tlNum
            mThreading.nextNum = nextResultStart
            'addLOG(qry + "," + nextResultStart.ToString)
            ThreadPool.QueueUserWorkItem(w, mThreading)
            Thread.Sleep(200)
            Application.DoEvents()
            Thread.Sleep(200)
            Application.DoEvents()
            nextResultStart += numResultsPerCall
        Loop

        numMS = DateDiff(DateInterval.Second, startReq, Now) * 1000

        Dim progressSoFar As Long
        progressSoFar = 0

        Do Until deviceS1.Count >= tlNum 'this counter will decrease based on AClient_searchWueryReceived event
            Application.DoEvents()
            Thread.Sleep(100)
            Application.DoEvents()
            progressSoFar += 100

            If progressSoFar Mod 1000 = 0 Then
                '1 second intervals
                addLOG("Still waiting.. # devices: " + deviceS1.Count.ToString)

                addLOG("Re-requesting " + deviceS1.Count.ToString + "-" + (mThreading.numRecs + deviceS1.Count).ToString)
                mThreading.nextNum = deviceS1.Count
                jsonText = AClient.searchAPI(mThreading)
                'this action kicks another request off if not received in a second
            End If
        Loop

allDone:

        addLOG(deviceS1.Count.ToString + " devices in " + Math.Round((progressSoFar + numMS) / 1000, 2).ToString("#.##") + " seconds")

        If queryBox1or2 = 1 Then
            Call enableButtons(1)
            QueryContainer1.btnGo.Visible = True
        Else
            Call enableButtons(2)
            QueryContainer2.btnGo.Visible = True
        End If

    End Sub

    Private Sub enableButtons(qry1or2 As Integer, Optional disableInstead As Boolean = False)
        Dim enablE As Boolean = True
        If disableInstead = True Then enablE = False

        If qry1or2 = 1 Then
            QueryContainer1.btnXls.Visible = enablE
            QueryContainer1.btnCsv.Visible = enablE
            QueryContainer1.btnJson.Visible = enablE
            '            QueryContainer1.btnGo.Visible = True
        Else
            QueryContainer2.btnXls.Visible = enablE
            QueryContainer2.btnCsv.Visible = enablE
            QueryContainer2.btnJson.Visible = enablE
            '           QueryContainer2.btnGo.Visible = True

        End If
    End Sub

    Private Sub addDevices(json$, ByRef allDevices As List(Of deviceData))
        ' query response handling
        Dim deviceResp As New List(Of deviceData)
        deviceResp = AClient.deserializeDeviceResponse(json)

        Dim numB4 As Long = allDevices.Count

        For Each D In deviceResp
            deviceS1.Add(D)
        Next

        processTasks1 = allDevices.Count

        'addLOG("Grew devices from " + numB4.ToString + " to " + allDevices.Count.ToString)
    End Sub
    Private Sub saveToClipboard(ByVal ttexT$)
        If IsNothing(ttexT) = True Or Len(ttexT) = 0 Then
            addLOG("Nothing to save to clipboard")
            Exit Sub
        End If

        My.Computer.Clipboard.SetText(ttexT)
        addLOG("JSON saved to clipboard")
    End Sub

    Private Sub QueryContainer1_exportJSON() Handles QueryContainer1.exportJSON
        Call saveToClipboard(jsoN1)
        MsgBox("JSON saved: " + Len(jsoN1).ToString + " chars")
    End Sub

    Private Sub AClient_searchQueryReturned(jsoN As String, firstNum As Long) Handles AClient.searchQueryReturned
        'addLOG("Devices starting at " + firstNum.ToString + " received")

        Dim a$
        If firstNum = 1200 Then
            a$ = ""
        End If

        Do Until deviceS1.Count = firstNum
            Thread.Sleep(200) ' event supposed to wait until it is its turn to add to deviceS
            Application.DoEvents()
        Loop

        Call addDevices(jsoN, deviceS1)

    End Sub
End Class
