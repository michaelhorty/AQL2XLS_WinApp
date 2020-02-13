Imports System.ComponentModel
Imports System.Threading
Imports System.Net


Public Class MainUI
    Private loggingEnabled As Boolean
    Private guiActive As Boolean
    Delegate Sub StringArgReturningVoidDelegate([text] As String)

    Public authInfo As apiAuthInfo
    Private activeToken$

    Private jsoN1$
    Private jsoN2$
    Private deviceS1 As List(Of deviceData)
    Private alertS1 As List(Of alertData)
    Private activitY1 As List(Of activityData)

    Private lockingObj As Boolean
    Private device1Nums As Collection

    Private WithEvents AClient As ARMclient

    Private ovaList As Collection
    Private deviceCompleted As Collection
    Private deviceString As Collection

    Private ovaFilename$
    Private Sub MainUI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' main entry point
        loggingEnabled = True
        guiActive = True
        Call multiControls()
        addLOG("GUI Activated")


        authInfo = New apiAuthInfo

        frmLogin.ShowDialog()

        '  With authInfo
        '      .fqdN = "https://" + TextBox1.Text + ".armis.com"
        '      .secretKey = TextBox2.Text
        '  End With

        Dim newClient As New ARMclient(authInfo)

        If Len(authInfo.tokeN) = 0 Then
            addLOG("Error getting token:" + newClient.lastError)
        Else
            'addLOG("Got token")
            addLOG("Token expires: UTC " + authInfo.tokenExpires.ToString("hh:mm:ss"))
            Call addLoginCreds(authInfo)
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
    Private Function getTArgs(qrY$, nNum As Long, nRecs As Integer) As threadingArgs
        getTArgs = New threadingArgs(authInfo)
        With getTArgs
            .nextNum = nNum
            .numRecs = nRecs
            .qrY = qrY
        End With
    End Function

    Private Sub submitQuery(qry$, Optional ByVal queryBox1or2 As Integer = 1)
        If authInfo.tokeN = "" Then
            MsgBox("Not authenticated.. Check your Secret Key!", vbOKOnly, "Not Connected")
            Exit Sub
        End If

        AClient = New ARMclient(authInfo)


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
        deviceString = New Collection
        deviceCompleted = New Collection
        device1Nums.Add(0)

        Dim mThreading As threadingArgs
        mThreading = getTArgs(qry, 0, numResultsPerCall)

        If mThreading.qType = "Undefined" Then
            MsgBox("SEARCH queries support Devices, Activities and Alerts. Please enter a valid query.", vbOKOnly, "Invalid Query")
            Exit Sub
        End If

        ' changing class globals didnt help
        '        AClient.currQry = qry
        '        AClient.currQueryType = mThreading.qType
        '        AClient.currBeginNDX = 0

        AClient.lastResult = ""


        Call AClient.searchAPI(mThreading)

        Do Until deviceCompleted.Count = 1
            Thread.Sleep(50)
        Loop

        If queryBox1or2 = 1 Then jsoN1 = jsonText Else jsoN2 = jsonText

        Dim respData As New deviceResponseData
        respData = AClient.deserializeResponseData(AClient.lastResult)

        tlNum = respData.total

        If tlNum <= 100 Then GoTo allDone

        Dim w As WaitCallback = New WaitCallback(AddressOf AClient.searchAPI)

        nextResultStart = respData.count

        'build coll first.. loop finishes queueing before source thread catches up
        Do Until nextResultStart >= tlNum
            device1Nums.Add(nextResultStart)
            nextResultStart += numResultsPerCall
        Loop

        addLOG("Queuing " + device1Nums.Count.ToString + " API hits for " + tlNum.ToString + " " + mThreading.qType)
        Dim numMS As Long = 0
        Dim startReq As DateTime = Now

        Dim numQ As Integer = 0

        lockingObj = False

        nextResultStart = numResultsPerCall

        ThreadPool.SetMaxThreads(20, 20)

        Do Until nextResultStart >= tlNum

            mThreading = getTArgs(qry, nextResultStart, numResultsPerCall)
            ThreadPool.QueueUserWorkItem(w, mThreading)
            numQ += 1
            QueryContainer1.Label1.Text = "Objects:" + (deviceCompleted.Count * 100).ToString
            Application.DoEvents()
            Thread.Sleep(150)
            Application.DoEvents()
            ' Do While numQ - deviceCompleted.Count > 20
            '     Thread.Sleep(50)
            '     Application.DoEvents()
            ' Loop
            nextResultStart += numResultsPerCall
        Loop

        numMS = DateDiff(DateInterval.Second, startReq, Now) * 1000

        Dim progressSoFar As Long
        progressSoFar = 0

        Dim numRetries As Integer = 0

        Do Until deviceCompleted.Count = device1Nums.Count 'this counter will decrease based on AClient_searchWueryReceived event
            Application.DoEvents()
            QueryContainer1.Label1.Text = "Objects:" + (deviceCompleted.Count * 100).ToString
            Thread.Sleep(100)
            Application.DoEvents()
            progressSoFar += 100
        Loop


allDone:


        deviceS1 = New List(Of deviceData)

        ' addLOG(deviceS1.Count.ToString + " objects being deserialized")


        Dim K As Integer = 0
        Dim wON As Integer = 0

        For Each JS In deviceString
            wON = grpNDX(deviceCompleted, K.ToString)
            If wON = 0 Then
                addLOG("Cannot find items starting at " + (K.ToString * 100).ToString)
            Else
                Call addDevices(JS, deviceS1, deviceCompleted(wON))
            End If
            K += 1
        Next

        QueryContainer1.Label1.Text = "[Objects]:" + deviceS1.Count.ToString

        addLOG(deviceS1.Count.ToString + " objects in " + Math.Round((progressSoFar + numMS) / 1000, 2).ToString("#.##") + " sec") 's (" + numRetries.ToString + " retries)")

        If queryBox1or2 = 1 Then
            Call enableButtons(1)
            QueryContainer1.btnGo.Visible = True
        Else
            Call enableButtons(2)
            QueryContainer2.btnGo.Visible = True
        End If

    End Sub

    Private Function getObjCount(mT As threadingArgs) As Long
        getObjCount = 0

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
        MsgBox("This feature is not fully implemented - currently you get only the most recent JSON response from the server (100 items per file)" + vbCrLf + "JSON saved to clipboard: " + Len(jsoN1).ToString + " chars", vbOKOnly, "JSON Export")
    End Sub

    Private Sub AClient_searchQueryReturned(jsoN As String, firstNum As Long, qryType As String) Handles AClient.searchQueryReturned
        'addLOG("Devices starting at " + firstNum.ToString + " received")


        If lockingObj = True Then
            Do Until lockingObj = False
                Application.DoEvents()
                Thread.Sleep(100)
            Loop
            '            lockingObj = True
        End If
        '
        lockingObj = True


        'If waitingOnNDX = 0 Then
        'addLOG("Received " + firstNum.ToString + " but did not expect it")
        'lockingObj = False
        'Exit Sub
        'Else
        'process search results
        'remove from queue

        Dim j$ = Mid(jsoN, InStr(jsoN, "next") - 1, 20)
        j = Mid(j, 1, InStr(j, ",") - 1)

        addLOG("Rcvd " + firstNum.ToString + ".. processing. JSON: " + j)
        '        threadResults(waitingOnNDX - 1) = jsoN
        deviceString.Add(jsoN)
        If firstNum > 0 Then firstNum = firstNum / 100
        deviceCompleted.Add(firstNum)
        'End If

        lockingObj = False
        Exit Sub

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

        Application.DoEvents()

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
                ReDim myXLS3d(deviceS1.Count, 16)

                addLOG("Creating Big 3D")
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
                    .Add("Site Name")
                    .Add("Site Location")
                    .Add("Sensor Name")
                    .Add("Visibility")
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
                        If IsNothing(.site) = False Then myXLS3d(roW, 13) = .site.name
                        If IsNothing(.site) = False Then myXLS3d(roW, 14) = .site.location
                        If IsNothing(.sensor) = False Then myXLS3d(roW, 15) = .sensor.name
                        If IsNothing(.visibility) = False Then myXLS3d(roW, 16) = .visibility
                        roW += 1
                    End With
                Next


                addLOG("Exporting-")

                Call dump2XLS(myXLS3d, roW, rArgs)
        End Select

    End Sub

    Private Sub QueryContainer1_MouseLeave(sender As Object, e As EventArgs) Handles QueryContainer1.MouseLeave

    End Sub

    Private Sub QueryContainer1_exportXLS() Handles QueryContainer1.exportXLS
        Dim rArgs As New reportingArgs
        rArgs.rptName = "Devices"
        rArgs.booL1 = True

        xlsEngine.RunWorkerAsync(rArgs)

    End Sub

    Private Sub xlsEngine_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles xlsEngine.RunWorkerCompleted
        addLOG("Export completed")
        GC.Collect()

    End Sub

    Private Sub QueryContainer1_exportCSV() Handles QueryContainer1.exportCSV
        Dim rArgs As New reportingArgs
        rArgs.rptName = "Devices"
        rArgs.booL1 = False

        Dim O As New SaveFileDialog
        O.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All files (*.*)|*.*"
        O.CheckPathExists = True
        O.Title = "Save"
        O.ShowDialog()

        rArgs.s1 = O.FileName

        xlsEngine.RunWorkerAsync(rArgs)

    End Sub

    Private Sub getOVA_DoWork(sender As Object, e As DoWorkEventArgs) Handles getOVA.DoWork
        If authInfo.tokeN = "" Then
            MsgBox("Not authenticated.. Check your Secret Key!", vbOKOnly, "Not Connected")
            Exit Sub
        End If
        Dim a$ = ""

        Dim ovaInfo As ovaArgs
        ovaInfo = e.Argument

        AClient = New ARMclient(authInfo)

        Dim currStatus$

        a$ = AClient.getImageStatus(authInfo, ovaInfo.ovaID.ToString)
        If InStr(a, "COMPLETE") Or InStr(a, "PROCESSI") Then
            addLOG("ALREADY EXISTS: Image <7 days old")
            '            MsgBox("Image requested <7 days ago." + vbCrLf + a, vbOKOnly, "Image already requested")
            currStatus = a
            GoTo downloadIMAGE
        End If

        a$ = AClient.makeImage(authInfo, ovaInfo)
        If a <> "True" Then
            addLOG("ERROR:" + a)
            Exit Sub
        Else
            addLOG("Submitted OVA request for collector #" + ovaInfo.ovaID.ToString)
        End If

        a$ = ""
        Thread.Sleep(3000)

        a$ = AClient.getImageStatus(authInfo, ovaInfo.ovaID.ToString)

        currStatus = AClient.deserializeImageStatus(a)

        If InStr(currStatus, "NOT_REQUESTED") Then
            addLOG("OVA" + ovaInfo.ovaID.ToString + ": NOT_REQUESTED-" + "There is a problem with the job. Contact support.")
            Exit Sub
        End If
        addLOG("OVA" + ovaInfo.ovaID.ToString + ": " + currStatus)

        Do Until InStr(currStatus, "COMPLETE")
            Application.DoEvents()
            Thread.Sleep(10000)
            Application.DoEvents()
            a$ = AClient.getImageStatus(authInfo, ovaInfo.ovaID.ToString)

            If Len(a) = 0 Then GoTo nextOne
            If InStr(a, "status") = 0 Then GoTo nextOne

            currStatus = AClient.deserializeImageStatus(a)

            If InStr(currStatus, "NOT_REQUESTED") Then
                addLOG("OVA" + ovaInfo.ovaID.ToString + ": NOT_REQUESTED-" + "There is a problem with the job. Contact support.")
                Exit Sub
            End If

            addLOG("OVA" + ovaInfo.ovaID.ToString + ": " + currStatus)

nextOne:
        Loop

        'a$ = "must have complete here now"

downloadIMAGE:
        Dim dlImg$
        dlImg = AClient.deserializeImageStatus(currStatus, True)

        ovaFilename = dlImg

        addLOG("OVA" + ovaInfo.ovaID.ToString + ": " + currStatus)

        'MsgBox("Image good for 7 days" + vbCrLf + a, vbOKOnly, "Image created")
    End Sub

    Private Sub btnOVA_Click(sender As Object, e As EventArgs) Handles btnOVA.Click

        If authInfo.tokeN = "" Then
            MsgBox("Not authenticated.. Check your Secret Key!", vbOKOnly, "Not Connected")
            Exit Sub
        End If
        Dim a$ = ""
        btnOVA.Visible = False

        AClient = New ARMclient(authInfo)

        a$ = AClient.getAvailOVAs(authInfo)

        If a = "False" Then
            addLOG("Unable to get OVA List")
            Exit Sub
            btnOVA.Visible = True
        End If

        Dim nL As List(Of availOVAresp)
        ovaList = AClient.deserializeOVAList(a)

        Call NewOVActl1.setupCTL(ovaList)

        QueryContainer1.Visible = False
        QueryContainer2.Visible = False
        NewOVActl1.Visible = True

        Exit Sub


    End Sub

    Private Sub NewOVActl1_userCancelled() Handles NewOVActl1.userCancelled
        NewOVActl1.Visible = False
        QueryContainer1.Visible = True
        addLOG("User cancelled OVA creation")
        btnOVA.Visible = True
    End Sub

    Private Sub NewOVActl1_createOVA(ipAddy As String, netMask As String, gateWay As String, ntpServer As String, dnsServer As String, proxy As String, ovaID As String) Handles NewOVActl1.createOVA
        NewOVActl1.Visible = False
        QueryContainer1.Visible = True

        Dim ovaDetails As ovaArgs
        ovaDetails = New ovaArgs

        addLOG("Creating new OVA")

        With ovaDetails
            .ip = ipAddy
            .netM = netMask
            .gWay = gateWay
            .ntP = ntpServer
            .dnS = dnsServer
            .proxY = proxy
            .ovaID = ovaID
        End With

        getOVA.RunWorkerAsync(ovaDetails)
    End Sub

    Private Sub getOVA_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles getOVA.RunWorkerCompleted
        If ovaFilename = "" Then GoTo nothingtoDo

        Dim O As New SaveFileDialog
        O.Filter = "OVA Files (*.ova)|*.ova|All files (*.*)|*.*"
        O.CheckPathExists = True
        O.Title = "Download Image - Save File As"
        O.ShowDialog()

        addLOG("Downloading file..")
        Dim remoteUri As String = ovaFilename
        Dim fileName As String = O.FileName

        Using client As New WebClient()

            client.DownloadFile(remoteUri, fileName)
        End Using

        addLOG("File downloaded and saved to " + O.FileName)
        MsgBox("Download complete:" + vbCrLf + O.FileName, vbOKOnly, "Finished")
nothingtoDo:
        btnOVA.Visible = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub PictureBox1_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Len(authInfo.tokeN) = 0 Then Exit Sub

        If Now.ToUniversalTime > authInfo.tokenExpires Then
            addLOG("Token refresh")

            Dim newClient As New ARMclient(authInfo)

            If Len(authInfo.tokeN) = 0 Then
                addLOG("Error getting token:" + newClient.lastError)
            Else
                ' addLOG("Got new token - " + authInfo.tokeN)
                addLOG("New Token expires: UTC " + authInfo.tokenExpires.ToString("hh:mm:ss"))
                'Call addLoginCreds(authInfo)
            End If


        End If

    End Sub
End Class
