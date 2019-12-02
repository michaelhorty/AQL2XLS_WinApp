Imports System.Threading

Public Class MainUI
    Private loggingEnabled As Boolean
    Private guiActive As Boolean
    Delegate Sub StringArgReturningVoidDelegate([text] As String)
    Private authInfo As apiAuthInfo

    Private jsoN1$
    Private jsoN2$
    Private deviceS1 As List(Of deviceData)
    Private alertS1 As List(Of alertData)
    Private activitY1 As List(Of activityData)

    Private lockingObj As Boolean
    Private device1Nums As Collection

    Private WithEvents AClient As ARMclient

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
        Dim mThreading As threadingArgs
        mThreading = New threadingArgs(authInfo)
        mThreading.qrY = qry

        If mThreading.qType = "Undefined" Then
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


        device1Nums = New Collection
        device1Nums.Add(0)

        jsonText = AClient.searchAPI(mThreading)

        If queryBox1or2 = 1 Then jsoN1 = jsonText Else jsoN2 = jsonText

        Dim respData As New deviceResponseData
        respData = AClient.deserializeResponseData(jsoN1)

        tlNum = respData.total

        If tlNum <= 100 Then GoTo allDone

        Dim w As WaitCallback = New WaitCallback(AddressOf AClient.searchAPI)


        'depending on reliability
        'may need to identify requests that werent answered
        'if queuing changed from 200 to 100ms, UI seems to not return responses

        nextResultStart = respData.count

        'build coll first.. loop finishes queueing before source thread catches up
        Do Until nextResultStart >= tlNum
            device1Nums.Add(nextResultStart)
            nextResultStart += numResultsPerCall
        Loop

        addLOG("Queuing " + device1Nums.Count.ToString + " API threads for " + tlNum.ToString + " " + mThreading.qType)

        nextResultStart = respData.count


        lockingObj = False

        ThreadPool.SetMaxThreads(25, 25)

        Dim numMS As Long = 0
        Dim startReq As DateTime = Now

        Do Until nextResultStart >= tlNum
            mThreading.nextNum = nextResultStart
            'addLOG(qry + "," + nextResultStart.ToString)
            ThreadPool.QueueUserWorkItem(w, mThreading)
            QueryContainer1.Label1.Text = "Objects:" + getObjCount(mThreading).ToString
            Application.DoEvents()
            Thread.Sleep(150)
            Application.DoEvents()
            'QueryContainer1.Label1.Text = "Objects:" + getObjCount(mThreading).ToString
            'Thread.Sleep(50)
            'Application.DoEvents()
            If nextResultStart Mod 10000 = 0 Then addLOG(Math.Round(nextResultStart / 100).ToString + " threads queued")
            nextResultStart += numResultsPerCall
        Loop

        numMS = DateDiff(DateInterval.Second, startReq, Now) * 1000

        Dim progressSoFar As Long
        progressSoFar = 0

        Dim numRetries As Integer = 0

        Do Until device1Nums.Count = 0 'this counter will decrease based on AClient_searchWueryReceived event
            Application.DoEvents()
            QueryContainer1.Label1.Text = "Objects:" + getObjCount(mThreading).ToString
            Thread.Sleep(100)
            Application.DoEvents()
            progressSoFar += 100

            'need to only request tose which are missing
            'build collection
            If numMS + progressSoFar > 3000 And progressSoFar Mod 500 = 0 Then
                '    '1 second intervals
                '   For Each D In device1Nums
                'addLOG("Re-requesting " + device1Nums(1).ToString + "-" + (CInt(device1Nums(1)) + numResultsPerCall).ToString)
                numRetries += 1
                mThreading.nextNum = device1Nums(1)
                jsonText = AClient.searchAPI(mThreading)
                Thread.Sleep(50)
                Application.DoEvents()
                '    Next
                '    'this action kicks another request off if not received in a second
            End If
        Loop

        QueryContainer1.Label1.Text = "[Objects]:" + getObjCount(mThreading).ToString

allDone:

        addLOG(getObjCount(mThreading).ToString + " objects in " + Math.Round((progressSoFar + numMS) / 1000, 2).ToString("#.##") + " secs (" + numRetries.ToString + " retries)")

        If queryBox1or2 = 1 Then
            Call enableButtons(1)
            QueryContainer1.btnGo.Visible = True
        Else
            Call enableButtons(2)
            QueryContainer2.btnGo.Visible = True
        End If

    End Sub

    Private Function getObjCount(mT As threadingArgs) As Long
        Select Case mT.qType
            Case "Devices"
                Return deviceS1.Count
            Case "Alerts"
                Return alertS1.Count
            Case "Activity"
                Return activitY1.Count

        End Select
    End Function

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

    Private Sub addActivity(json$, ByRef allActivities As List(Of activityData), indeX As Long)
        ' query response handling
        Dim deviceResp As New List(Of activityData)
        deviceResp = AClient.deserializeActivityResponse(json)

        Dim numB4 As Long = allActivities.Count
        Dim K As Integer = 0



        If indeX > allActivities.Count Then
            For K = 0 To deviceResp.Count - 1
                allActivities.Add(deviceResp(K))
            Next

        Else
            For K = 0 To deviceResp.Count - 1
                allActivities.Insert(indeX + K, deviceResp(K))
            Next
        End If


        'addLOG("Grew activities from " + numB4.ToString + " to " + allActivities.Count.ToString)
    End Sub

    Private Sub addAlerts(json$, ByRef allAlerts As List(Of alertData), index As Long)
        ' query response handling
        Dim deviceResp As New List(Of alertData)
        deviceResp = AClient.deserializeAlertResponse(json)

        Dim K As Integer = 0
        Dim numB4 As Long = allAlerts.Count


        If index > allAlerts.Count Then
            For K = 0 To deviceResp.Count - 1
                allAlerts.Add(deviceResp(K))
            Next

        Else
            For K = 0 To deviceResp.Count - 1
                allAlerts.Insert(index + K, deviceResp(K))
            Next
        End If


        'addLOG("Grew alerts from " + numB4.ToString + " to " + allAlerts.Count.ToString)
    End Sub

    Private Sub addDevices(json$, ByRef allDevices As List(Of deviceData), indeX As Long)
        ' query response handling


        Dim deviceResp As New List(Of deviceData)
        deviceResp = AClient.deserializeDeviceResponse(json)

        Dim numB4 As Long = allDevices.Count

        Dim K As Integer

        If indeX > allDevices.Count Then
            For K = 0 To deviceResp.Count - 1
                allDevices.Add(deviceResp(K))
            Next

        Else
            For K = 0 To deviceResp.Count - 1
                allDevices.Insert(indeX + K, deviceResp(K))
            Next
        End If


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

    Private Sub AClient_searchQueryReturned(jsoN As String, firstNum As Long, qryType As String) Handles AClient.searchQueryReturned
        'addLOG("Devices starting at " + firstNum.ToString + " received")


        If lockingObj = True Then
            Do Until lockingObj = False
                Application.DoEvents()
                Thread.Sleep(100)
            Loop
            lockingObj = True
        End If

        Dim waitingOnNDX = grpNDX(device1Nums, firstNum.ToString)

        If waitingOnNDX = 0 Then
            'addLOG("Received " + firstNum.ToString + " but did not expect it")
            lockingObj = False
            Exit Sub
        Else
            'process search results
            'remove from queue

            'addLOG("Received " + firstNum.ToString + ".. processing")
            device1Nums.Remove(waitingOnNDX)
        End If

        Select Case qryType
            Case "Devices"
                If firstNum = 0 Then deviceS1 = New List(Of deviceData)

                'Do Until deviceS1.Count = firstNum
                '    Thread.Sleep(200) ' event supposed to wait until it is its turn to add to deviceS
                '    Application.DoEvents()
                'Loop

                Call addDevices(jsoN, deviceS1, firstNum)

            Case "Activity"
                If firstNum = 0 Then activitY1 = New List(Of activityData)
                '        Do Until activitY1.Count = firstNum
                '             Thread.Sleep(200) ' event supposed to wait until it is its turn to add to deviceS
                '              Application.DoEvents()
                '           Loop

                Call addActivity(jsoN, activitY1, firstNum)

            Case "Alerts"
                If firstNum = 0 Then alertS1 = New List(Of alertData)
                '             Do Until alertS1.Count = firstNum
                '                  Thread.Sleep(200) ' event supposed to wait until it is its turn to add to deviceS
                '                   Application.DoEvents()
                '                Loop

                Call addAlerts(jsoN, alertS1, firstNum)

        End Select

        lockingObj = False


    End Sub

    Private Sub xlsEngine_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles xlsEngine.DoWork
        Dim rArgs As reportingArgs = e.Argument

        Dim rptType$ = rArgs.rptName

        Dim myXLS3d(1000000, 50) As Object

        Dim coL As Long = 1
        Dim roW As Long = 0

        Select Case rArgs.rptName
            Case "Devices"
                'For Each C In rArgs.someColl

                'coL += 1
                'Next
                ReDim myXLS3d(deviceS1.Count - 1, 12)

                rArgs.someColl = New Collection
                With rArgs.someColl
                    .Add("Category")
                    .Add("MacAddress")
                    .Add("Name")
                    .Add("Manufacturer")
                    .Add("Model")
                    .Add("OS")
                    .Add("OS Version")
                    .Add("User")
                    .Add("IP")
                    .Add("FirstSeen")
                    .Add("LastSeen")
                    .Add("ID")
                    .Add("RiskLevel")
                End With

                For Each D In deviceS1
                    With D
                        If IsNothing(.category) = False Then myXLS3d(roW, 0) = .category

                        If IsNothing(.macAddress) = False Then myXLS3d(roW, 1) = .macAddress.ToString
                        If IsNothing(.name) = False Then myXLS3d(roW, 2) = .name
                        If IsNothing(.manufacturer) = False Then myXLS3d(roW, 3) = .manufacturer
                        If IsNothing(.model) = False Then myXLS3d(roW, 4) = .model
                        If IsNothing(.operatingSystem) = False Then myXLS3d(roW, 5) = .operatingSystem
                        If IsNothing(.operatingSystemVersion) = False Then myXLS3d(roW, 6) = .operatingSystemVersion.ToString
                        If IsNothing(.user) = False Then myXLS3d(roW, 7) = .user

                        If IsNothing(.ipAddress) = False Then myXLS3d(roW, 8) = .ipAddress.ToString
                        If IsNothing(.firstSeen) = False Then myXLS3d(roW, 9) = .firstSeen.ToString
                        If IsNothing(.lastSeen) = False Then myXLS3d(roW, 10) = .lastSeen.ToString
                        If IsNothing(.id) = False Then myXLS3d(roW, 11) = .id.ToString
                        If IsNothing(.riskLevel) = False Then myXLS3d(roW, 12) = .riskLevel.ToString
                        roW += 1
                    End With
                Next

                rArgs.booL1 = True

                Call dump2XLS(myXLS3d, roW, rArgs)
        End Select

    End Sub

    Private Sub QueryContainer1_MouseLeave(sender As Object, e As EventArgs) Handles QueryContainer1.MouseLeave

    End Sub

    Private Sub QueryContainer1_exportXLS() Handles QueryContainer1.exportXLS
        Dim rArgs As New reportingArgs
        rArgs.rptName = "Devices"

        xlsEngine.RunWorkerAsync(rArgs)

    End Sub
End Class
