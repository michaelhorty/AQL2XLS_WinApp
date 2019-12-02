Imports RestSharp
Imports Newtonsoft.Json



Public Class threadingArgs
    Public authInfo As apiAuthInfo
    Public nextNum As Long
    Public numRecs As Long
    Public qrY$

    Public ReadOnly Property qType As String
        Get
            Dim objType1$ = "Undefined"
            If LCase(Mid(qrY, 1, 10)) = "in:devices" Then objType1 = "Devices"
            If LCase(Mid(qrY, 1, 13)) = "in:activity" Then objType1 = "Activity"
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
    Public tags$() ' As deviceTags
    Public type$
    Public user$

End Class

Public Class alertData

End Class

Public Class sensorData
    Public name$
End Class
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

    Public Event searchQueryReturned(jsoN$, firstNum As Long, qryType As String)
    Public Sub New(ByRef authInfo As apiAuthInfo)
        'when setting up new client
        'can use token from previous clients
        'check expiry in here to obtain a new secret_key whenever necessary
        'this way can set up clients on the fly (multithreading)

        If Len(authInfo.tokeN) Then
            'already authorized, no need to get new token
            With authInfo
                fqdN = .fqdN
                tokenExpires = .tokenExpires
                tokeN = .tokeN
            End With
            Exit Sub
        End If

        fqdN = authInfo.fqdN
        theKey = authInfo.secretKey

        gotToken = False

        Call GetToken(authInfo)
    End Sub


    Private Sub setError(ByVal theError$)
        MainUI.addLOG("ERROR: " + theError)
        lastError = Now.ToLongTimeString + " - " + theError
    End Sub


    Private Function GetToken(ByRef authInfo As apiAuthInfo) As String
        'On Error GoTo errorcatch
        Dim secretKey$ = authInfo.secretKey

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

    Public Function deserializeDeviceResponse(ByRef json$) As List(Of deviceData)
        deserializeDeviceResponse = New List(Of deviceData)

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

        deserializeDeviceResponse = JsonConvert.DeserializeObject(Of List(Of deviceData))(jList)
    End Function

    Public Function deserializeAlertResponse(ByRef json$) As List(Of AlertData)
        deserializeAlertResponse = New List(Of AlertData)

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

        deserializeAlertResponse = JsonConvert.DeserializeObject(Of List(Of AlertData))(jList)
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

    Public Function searchAPI(tInfo As threadingArgs) As String
        'test for expired token & get new if necessary
        If DateDiff(DateInterval.Minute, Now.ToUniversalTime, tokenExpires) < 5 Then
            'within 5 minutes of being expired, get new token
            Call GetToken(tInfo.authInfo)
        End If

        searchAPI = ""

        Dim qryString$ = fqdN + "/api/v1/search/?aql=" + tInfo.qrY + "&from=" + tInfo.nextNum.ToString + "&length=" + tInfo.numRecs.ToString

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

        succesS = CBool(getJSONObject("success", a))

        If response.IsSuccessful = False Or succesS = False Then
            Call setError("Error to " + qryString + " - " + msgResp)
        End If

        If succesS = True Then ' if api response shows true
            'MainUI.addLOG("Search query returned - " + qryString)
        End If

        searchAPI = a

        RaiseEvent searchQueryReturned(a, tInfo.nextNum, tInfo.qType)

    End Function

End Class
