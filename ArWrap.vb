﻿Imports RestSharp
Imports Newtonsoft.Json


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

        If CBool(GetToken(theKey)) = True Then
            gotToken = True
            authInfo.tokeN = tokeN
            authInfo.tokenExpires = tokenExpires
        End If
    End Sub

    Private Sub newToken(setNewToken$)
        gotToken = True
        tokeN = tokeN
    End Sub

    Private Sub setError(ByVal theError$)
        MainUI.addLOG("ERROR: " + theError)
        lastError = Now.ToLongTimeString + " - " + theError
    End Sub


    Private Function GetToken(ByVal secretKey) As String
        'On Error GoTo errorcatch
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

        Dim a$ = "Response empty"
        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean
        succesS = CBool(getJSONObject("success", a))

        tokeN$ = ""

        If response.IsSuccessful = False Then
            Call setError("Did Not get token response - " + msgResp)
            Return a
            Exit Function
        End If

        Dim tokData As New tokenContainer
        tokData = JsonConvert.DeserializeObject(Of tokenContainer)(response.Content)
        tokeN = tokData.data.access_token
        tokenExpires = CDate(tokData.data.expiration_utc)

        Return "True"
        Exit Function
errorcatch:
        Call setError("Token error - could Not make request:   " + a) ' " + response.ErrorMessage.ToString + " " + response.Content)

        Return ""
    End Function

    Public Function getJSONObject(key$, json$) As String
        On Error GoTo errorcatch

        Dim sObject = JsonConvert.DeserializeObject(json)
        Return sObject(key)

errorcatch:
        Return ""
    End Function

    Public Function searchAPI(ByVal qrY$, Optional ByVal beginNum As Long = 0, Optional ByVal endNum As Long = 100) As String
        Dim client = New RestClient(fqdN + "/api/v1/search/?aql=" + qrY)

        Dim request = New RestRequest(Method.GET)
        request.AddHeader("Cache-Control", "no-cache")
        request.AddHeader("Authorization", tokeN)

        Dim response As IRestResponse
        response = client.Execute(request)
        ' age"": ""Invalid access token.""," & vbLf & "  ""success
        Dim a$ = "Response empty"
        If IsNothing(response.Content) = False Then a$ = response.Content
        If a = "" Then a = response.ErrorMessage

        Dim msgResp$
        msgResp = getJSONObject("message", a)
        Dim succesS As Boolean
        succesS = CBool(getJSONObject("success", a))

        tokeN$ = ""

        If response.IsSuccessful = False Then
            Call setError("Did Not get token response - " + msgResp)
            Return a
            Exit Function
        End If

    End Function

End Class

Public Class ArWrapper

    Public apiURI$
    Public apiResponse$
    Public rsExecuted As Boolean
    Private rsRequest As RestSharp.RestRequest
    Private getORpost$
    Private restToken$

    Public Function exeAPI() As IRestResponse
        rsExecuted = False
        Dim rsClient As RestClient

        rsClient = New RestClient(Me.apiURI)

        Me.rsExecuted = True
        exeAPI = rsClient.Execute(rsRequest)

        Me.apiResponse = exeAPI.Content
        MainUI.addLOG("-------ARM REQUEST--------")
        MainUI.addLOG(Me.getORpost + ": " + Me.apiURI)
        Dim c$ = ""

        For Each P In rsRequest.Parameters
            c = P.Value.ToString
            If Mid(c, 1, 6) = "Authorization" Then c = "Authorization {token}"
            MainUI.addLOG(P.Name.ToString + " " + c)
        Next
        MainUI.addLOG("--------------------------")

        If Len(Me.apiResponse) = 0 Then
            Me.apiResponse = exeAPI.ErrorMessage
        End If

    End Function

End Class
