Imports RestSharp
Imports Newtonsoft.Json

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
