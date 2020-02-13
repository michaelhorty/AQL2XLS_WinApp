Imports RestSharp
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Public Class ovaArgs
    Public ip$
    Public netM$
    Public gWay$
    Public ntP$
    Public dnS$
    Public proxY$
    Public ovaID$
    Public saveFilename$
End Class

Public Class threadingArgs
    Public authInfo As apiAuthInfo
    Public nextNum As Long
    Public numRecs As Long
    Public qrY$

    Public ReadOnly Property qType As String
        Get
            Dim objType1$ = "Undefined"
            If LCase(Mid(qrY, 1, 10)) = "in:devices" Then objType1 = "Devices"
            If LCase(Mid(qrY, 1, 11)) = "in:activity" Then objType1 = "Activity"
            If LCase(Mid(qrY, 1, 9)) = "in:alerts" Then objType1 = "Alerts"
            Return objType1
        End Get

    End Property

    Public Sub New(A As apiAuthInfo)
        authInfo = A
        nextNum = 0
        numRecs = 100
    End Sub

End Class
Public Class deviceResponseData
    Public count As Long
    'Public Next As Long 'next is lang keyword cannot set json
    Public prev As Long
    Public nextResults As Long
    Public total As Long

End Class
Public Class sensorData
    Public name$
End Class
Public Class siteData
    Public location$
    Public name$
End Class
Public Class deviceData
    Public category$
    Public firstSeen$
    Public id As Long
    Public ipAddress$
    Public lastSeen$
    Public macAddress$
    Public manufacturer$
    Public model$
    Public name$
    Public operatingSystem$
    Public operatingSystemVersion$
    Public riskLevel As Long
    Public sensor As sensorData
    Public site As siteData
    Public tags$() ' As deviceTags
    Public type$
    Public user$
    Public visibility$

End Class
Public Class availOVAresp
    Public ova_ids() As String
End Class
Public Class alertData

End Class

'Public Class sensorData'
'Public name '$
'End Class
Public Class activityData
    Public activityId$
    Public activityUUID$
    Public connectionIds$()
    Public content$
    Public deviceIds() As Long
    Public protocol$
    Public sensor As sensorData
    Public time$
    Public title$
    Public type$

End Class
Public Class tokenData
    Public access_token As String
    Public expiration_utc As String
End Class

Public Class tokenContainer
    Public data As tokenData
End Class

Public Class apiAuthInfo
    Public tokeN$
    Public fqdN$
    Public tokenExpires As DateTime
    Public secretKey$
End Class

Public Class ARMclient
    Public lastError$ = ""
    Public tokeN$ 'token used for REST calls
    Public fqdN$ 'ie http://localhost https://myserver.myzone.com
    Public gotToken As Boolean
    Public tokenExpires As DateTime
    Private theKey$

    Private currCreds As apiAuthInfo

    Public currQry$
    Public currBeginNDX As Long
    Public currQueryType$

    Public lastResult$

    Public Event searchQueryReturned(ByVal jsoN$, ByVal firstNum As Long, ByVal qryType As String)
    Public Sub New(ByRef authInfo As apiAuthInfo)
        'when setting up new client
        'can use token from previous clients
        'check expiry in here to obtain a new secret_key whenever necessary
        'this way can set up clients on the fly (multithreading)

        If Len(authInfo.tokeN) And Now.ToUniversalTime < authInfo.tokenExpires Then
            'already authorized, no need to get new token
            With authInfo
                fqdN = .fqdN
                tokenExpires = .tokenExpires
                tokeN = .tokeN
            End With
            Exit Sub
        End If

        currCreds = New apiAuthInfo

        With currCreds
            .secretKey = authInfo.secretKey
            .fqdN = authInfo.fqdN
            fqdN = .fqdN
            theKey = .secretKey
        End With

        gotToken = False

        Call GetToken(currCreds)

        With authInfo
            .secretKey = currCreds.secretKey
            .tokeN = currCreds.tokeN
            .tokenExpires = currCreds.tokenExpires
            tokeN = .tokeN
            tokenExpires = .tokenExpires
        End With
    End Sub


    Private Sub setError(ByVal theError$)
        MainUI.addLOG("ERROR: " + theError)
        lastError = Now.ToLongTimeString + " - " + theError
    End Sub


    Private Function GetToken(ByRef authInfo As apiAuthInfo) As String
        'On Error GoTo errorcatch
        Dim secretKey$ = authInfo.secretKey

        If Len(theKey) Then secretKey = theKey



        Dim a$ = "Response empty"
        GetToken = a

        If Len(secretKey) = 0 Then
            Exit Function
        End If

        Dim client = New RestClient(fqdN + "/api/v1/access_token/")
        Dim request = New RestRequest(Method.POST)
        request.AddHeader("Cache-Control", "no-cache")
        request.AddHeader("Content-Type", "application/x-www-form-urlencoded")
        '        request.AddParameter("multipart / Form - Data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW", "------WebKitFormBoundary7MA4YWxkTrZu0gW" + Chr(13) + "Content-Disposition: Form-Data; name=" + Chr(34) + "secret_key" + Chr(34) + Chr(13) + secretKey + Chr(13) + "------WebKitFormBoundary7MA4YWxkTrZu0gW--", ParameterType.RequestBody)
        '        request.AddParameter("undefined", "secret_key=" + secretKey, ParameterType.RequestBody)
        '
        request.AddParameter("undefined", "secret_key=" + secretKey, ParameterType.RequestBody)

        Dim response As IRestResponse
        response = client.Execute(request)

        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean
        succesS = CBool(getJSONObject("success", a))

        tokeN$ = ""

        If response.IsSuccessful = False Then
            Call setError("Did Not get token response - " + msgResp)
            Return "False"
            Exit Function
        End If

        Dim tokData As New tokenContainer
        tokData = JsonConvert.DeserializeObject(Of tokenContainer)(response.Content)
        tokeN = tokData.data.access_token
        tokenExpires = CDate(tokData.data.expiration_utc)

        authInfo.tokeN = tokeN
        authInfo.tokenExpires = tokenExpires

        Return "True"
        Exit Function

    End Function

    Public Function deserializeOVAList(ByRef json$) As Collection
        Dim jList$ = ""
        jList = Mid(json, InStr(json, "[") + 1)
        jList = Mid(jList, 1, InStr(jList, "]") - 1)
        jList = Replace(jList, " ", "")
        jList = Replace(jList, vbLf, "")

        Dim K As Long
        Dim a As Object
        a = Split(jList, ",")

        deserializeOVAList = New Collection

        For K = 0 To UBound(a)
            deserializeOVAList.Add(a(K))
        Next


    End Function

    Public Function deserializeDeviceResponse(ByRef json$) As List(Of deviceData)
        'On Error Resume Next

        deserializeDeviceResponse = New List(Of deviceData)

        Dim jList$ = ""
        jList = Mid(json, InStr(json, "[") - 1)
        If Mid(jList, 1, 1) = ":" Then jList = Mid(jList, 2)

        Dim K As Long
        Dim a$ = ""

        K = Len(jList)
        a = Mid(jList, K, 1)

        Do Until a = "]"
            K = K - 1
            jList = Mid(jList, 1, K)
            a = Mid(jList, K, 1)
        Loop

        deserializeDeviceResponse = JsonConvert.DeserializeObject(Of List(Of deviceData))(jList)
    End Function

    Public Function deserializeImageStatus(ByRef jsoN$, Optional ByVal downloadURLonly As Boolean = False) As String
        If Mid(jsoN, 1, 14) = "COMPLETE LINK:" Then
            deserializeImageStatus = Mid(jsoN, 15)
            Exit Function
        End If

        Dim jsonObject As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(jsoN)
        '        Dim jsonArray As JArray = jsonObject("result")

        deserializeImageStatus = ""
        deserializeImageStatus += jsonObject("data").Item("status").ToString

        If InStr(deserializeImageStatus, "COMPLETE") Then
            deserializeImageStatus += " LINK:" + jsonObject("data").Item("downloadUrl").ToString
            If downloadURLonly = True Then deserializeImageStatus = jsonObject("data").Item("downloadUrl").ToString
        End If

        '.count = jsonObject("data").Item("count")
        '    .prev = jsonObject("data").Item("prev")
        '    .nextResults = jsonObject("data").Item("next")
        'End With


    End Function

    Public Function deserializeAlertResponse(ByRef json$) As List(Of alertData)
        deserializeAlertResponse = New List(Of alertData)

        Dim jList$ = ""
        jList = Mid(json, InStr(json, "[") - 1)

        Dim K As Long
        Dim a$ = ""

        K = Len(jList)
        a = Mid(jList, K, 1)

        Do Until a = "]"
            K = K - 1
            jList = Mid(jList, 1, K)
            a = Mid(jList, K, 1)
        Loop

        deserializeAlertResponse = JsonConvert.DeserializeObject(Of List(Of alertData))(jList)
    End Function

    Public Function deserializeActivityResponse(ByRef json$) As List(Of ActivityData)
        deserializeActivityResponse = New List(Of ActivityData)

        Dim jList$ = ""
        jList = Mid(json, InStr(json, "[") - 1)

        Dim K As Long
        Dim a$ = ""

        K = Len(jList)
        a = Mid(jList, K, 1)

        Do Until a = "]"
            K = K - 1
            jList = Mid(jList, 1, K)
            a = Mid(jList, K, 1)
        Loop

        deserializeActivityResponse = JsonConvert.DeserializeObject(Of List(Of ActivityData))(jList)
    End Function


    Public Function deserializeResponseData(ByRef json$) As deviceResponseData
        On Error Resume Next

        deserializeResponseData = New deviceResponseData

        If Len(json) = 0 Then Exit Function


        Dim jsonObject As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(json)
        '        Dim jsonArray As JArray = jsonObject("result")

        With deserializeResponseData
            .total = jsonObject("data").Item("total")
            .count = jsonObject("data").Item("count")

            .prev = jsonObject("data").Item("prev")
            .nextResults = jsonObject("data").Item("next")
        End With

        Dim a$
        a$ = ""

    End Function
    Public Function getJSONObject(key$, json$) As String
        On Error GoTo errorcatch

        Dim sObject = JsonConvert.DeserializeObject(json)
        Return sObject(key)

errorcatch:
        Return ""
    End Function

    Public Function getAvailOVAs(ByRef authInfo As apiAuthInfo) As String
        If DateDiff(DateInterval.Minute, Now.ToUniversalTime, authInfo.tokenExpires) < 5 Then
            Call GetToken(authInfo)
        End If

        Dim qryString$ = fqdN + "/api/v1/ovas/" '

        Dim client = New RestClient(qryString)

        Dim request = New RestRequest(Method.GET)
        request.AddHeader("Cache-Control", "no-cache")
        request.AddHeader("Authorization", tokeN)

        Dim response As IRestResponse
        response = client.Execute(request)

        Dim a$ = "Response empty"


        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        If response.IsSuccessful = False Then
            Return a
        End If

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean

        succesS = CBool(getJSONObject("success", a))

        If succesS = False Then
            Return "False: " + msgResp
        Else
            Return a
        End If

    End Function
    Public Function getImageStatus(authInfo As apiAuthInfo, ovaID$) As String
        If DateDiff(DateInterval.Minute, Now.ToUniversalTime, authInfo.tokenExpires) < 5 Then
            Call GetToken(authInfo)
        End If

        Dim client = New RestClient(fqdN + "/api/v1/ova_create/" + ovaID + "/")
        Dim request = New RestRequest(Method.GET)
        request.AddHeader("Cache-Control", "no-cache")
        request.AddHeader("Authorization", tokeN)

        Dim response As IRestResponse
        response = client.Execute(request)

        Dim a$ = ""

        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        If response.IsSuccessful = False Then
            Return a
            Exit Function

        End If

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean
        succesS = CBool(getJSONObject("success", a))

        If succesS = False Then
            'MainUI.addLOG("ERROR:" + msgResp)
            Return msgResp
            Exit Function
        Else
            getImageStatus = a
        End If

        a$ = ""


    End Function
    Public Function makeImage(ByRef authInfo As apiAuthInfo, ByRef ovaInfo As ovaArgs) As String
        If DateDiff(DateInterval.Minute, Now.ToUniversalTime, authInfo.tokenExpires) < 5 Then
            Call GetToken(authInfo)
        End If

        Dim client = New RestClient(fqdN + "/api/v1/ova_create/" + ovaInfo.ovaID + "/")
        Dim request = New RestRequest(Method.POST)
        request.AddHeader("Cache-Control", "no-cache")
        request.AddHeader("Authorization", tokeN)

        'request.AddHeader("Content-Type", "application/x-www-form-urlencoded")

        request.AddHeader("Content-Type", "application/json")
        '        request.AddParameter("multipart / Form - Data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW", "------WebKitFormBoundary7MA4YWxkTrZu0gW" + Chr(13) + "Content-Disposition: Form-Data; name=" + Chr(34) + "secret_key" + Chr(34) + Chr(13) + secretKey + Chr(13) + "------WebKitFormBoundary7MA4YWxkTrZu0gW--", ParameterType.RequestBody)
        '        request.AddParameter("undefined", "secret_key=" + secretKey, ParameterType.RequestBody)
        '
        'request.AddParameter("undefined", "address=192.168.2.100", ParameterType.RequestBody) '  + secretKey, ParameterType.RequestBody)
        'request.AddParameter("undefined", "dnsNameservers=8.8.8.8,8.8.4.4", ParameterType.RequestBody)
        ''request.AddParameter("undefined", "gateway=192.168.0.1", ParameterType.RequestBody)
        'request.AddParameter("undefined", "netmask=255.255.0.0", ParameterType.RequestBody)
        'request.AddParameter("undefined", "ntpServers=1.2.3.4,my.ntp.com", ParameterType.RequestBody)

        Dim c$ = Chr(34)

        Dim bodY$ = ""

        'bodY = "{" + c$ + "address" + c$ + ":  " + c$ + ovainfo.ip + c$ + "," + c$ + "dnsNameservers" + c$ + ": [" + c$ + "8.8.8.8" + c$ + ", " + c$ + "8.8.4.4" + c$ + "]," + c$ + "gateway" + c$ + ": " + c$ + "192.168.0.1" + c$ + "," + c$ + "httpsProxy" + c$ + ": " + c$ + "https://myproxy.com" + c$ + "," + c$ + "netmask" + c$ + ": " + c$ + "255.255.0.0" + c$ + "," + c$ + "ntpServers" + c$ + ": [" + c$ + "1.2.3.4" + c$ + ", " + c$ + "my.ntp.com" + c$ + "]}"

        bodY = "{" + c$ + "address" + c$ + ":  " + c$ + ovaInfo.ip + c$ + "," + c$ + "dnsNameservers" + c$ + ": [" + csvTOquotedList(ovaInfo.dnS) + "]," + c$ + "gateway" + c$ + ": " + c$ + ovaInfo.gWay + c$ + "," + c$ + "netmask" + c$ + ": " + c$ + ovaInfo.netM + c$ + "," + c$ + "ntpServers" + c$ + ": [" + csvTOquotedList(ovaInfo.ntP) + "]}"

        request.AddJsonBody(bodY)

        Dim response As IRestResponse
        response = client.Execute(request)

        Dim a$ = ""

        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        If response.IsSuccessful = False Then
            Return a
            Exit Function

        End If

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean
        succesS = CBool(getJSONObject("success", a))

        If succesS = False Then
            'MainUI.addLOG("ERROR:" + msgResp)
            Return msgResp
            Exit Function
        Else
            makeImage = "True"
        End If

        a$ = ""


    End Function


    Public Sub searchAPI(qArgs As threadingArgs) ' As String
        'test for expired token & get new if necessary
        If DateDiff(DateInterval.Minute, Now.ToUniversalTime, qArgs.authInfo.tokenExpires) < 5 Then
            'within 5 minutes of being expired, get new token
            Call GetToken(qArgs.authInfo)
        End If

        'searchAPI = ""

        'Dim qrY$ = currQry
        'Dim currNDX As Long = currBeginNDX
        'Dim qryType$ = currQry

        Dim retrieS As Integer = 0

doOver:
        retrieS += 1

        If retrieS > 3 Then
            Call setError("Could not obtain new token - connection closed")
            lastResult = "Connection closed by server - could not refresh token"
            RaiseEvent searchQueryReturned(lastResult, qArgs.nextNum, "Devices")
            Exit Sub
        End If

        Dim qryString$ = fqdN + "/api/v1/search/?aql=" + qArgs.qrY + "&from=" + qArgs.nextNum.ToString + "&length=" + qArgs.numRecs.ToString

        Dim client = New RestClient(qryString)

        Dim request = New RestRequest(Method.GET)
        request.AddHeader("Cache-Control", "no-cache")
        request.AddHeader("Authorization", tokeN)

        Dim response As IRestResponse
        response = client.Execute(request)

        Dim a$ = "Response empty"


        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean

        Dim b$ = ""

        b = getJSONObject("success", a)

        If Len(b) Then succesS = CBool(b) Else succesS = False

        If response.IsSuccessful = False Or succesS = False Then
            If msgResp.ToString = "Invalid access token." Then
                Call GetToken(qArgs.authInfo)
                GoTo doOver
            End If
            Call setError("Error to " + qryString.ToString + " - " + msgResp.ToString + " / " + response.ErrorMessage.ToString)
        End If

        If succesS = True Then ' if api response shows true
            'MainUI.addLOG("Search query returned - " + qryString)
        End If

        'searchAPI = a

        lastResult = a

        RaiseEvent searchQueryReturned(a, qArgs.nextNum, "Devices")

    End Sub

End Class
