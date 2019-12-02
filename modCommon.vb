Imports System.Security.Cryptography
Imports System.IO

Module modCommon

    Public Sub saveJSONtoFile(jsonString$, ByVal errFN$, ByRef add2zip As Collection)

        Dim fileN$ = errFN
        '        If fileN = "" Then
        '        Dim CD As New SaveFileDialog
        '        CD.Title = "Select Save Filename"
        '        If Len(API.shortName) Then CD.Title += " for " + API.shortName
        '        CD.FileName = errFN
        '        CD.Filter = "TEXT files (*.txt)|*.txt|JSON files (*.json)|*.json|All files (*.*)|*.*"
        '        CD.ShowDialog()
        '        If Len(CD.FileName) Then fileN = CD.FileName
        '        End If

        '   If Len(fileN) = 0 Then
        '   MsgBox("Must enter a filename.", vbOKOnly, "Error")
        '   Exit Sub
        '   End If

        Call safeKILL(fileN)
        Call streamWriterTxt(fileN, jsonString)

        add2zip.Add(fileN)

        GC.Collect()

    End Sub

    Public Function beginWeek(ByVal someD As Date, Optional ByVal showEndInstead As Boolean = False, Optional ByVal defaultFirstDay As DayOfWeek = DayOfWeek.Monday) As Date
        Dim moveVal As Integer = -1

        Do Until someD.DayOfWeek = DayOfWeek.Monday
            someD = DateAdd(DateInterval.Day, moveVal, someD)
        Loop

        If showEndInstead = True Then
            someD = DateAdd(DateInterval.Day, 6, someD)
        End If
        Return someD
    End Function

    Public Function weekRange(ByVal someD As Date) As String
        Dim b As Date = beginWeek(someD)
        Dim e As Date = beginWeek(someD, True)
        weekRange = b.Month.ToString + "/" + b.Day.ToString + "-" + e.Month.ToString + "/" + e.Day.ToString
    End Function

    Public Function streamWriterTxt(fileN$, string2write$) As Boolean
        streamWriterTxt = True
        On Error GoTo errorCatch
        Dim fS As New FileStream(fileN, FileMode.OpenOrCreate, FileAccess.Write)
        Dim sW As New StreamWriter(fS)
        sW.BaseStream.Seek(0, SeekOrigin.End)
        sW.WriteLine(string2write$)
        sW.Flush()
        sW.Close()

        sW = Nothing
        fS = Nothing

        Exit Function

errorCatch:
        streamWriterTxt = False
        sW = Nothing
        fS = Nothing


    End Function

    Public Function grpNDX(ByRef C As Collection, ByRef a$, Optional ByVal caseSensitive As Boolean = True) As Integer
        Dim K As Long
        grpNDX = 0

        If C.Count = 0 Then Exit Function
        If caseSensitive = False Then GoTo dontEvalCase

        For K = 1 To C.Count
            If a = C(K) Then
                grpNDX = K
                Exit Function
            End If
        Next
        Exit Function

dontEvalCase:
        For Each S In C
            K += 1
            If LCase(a) = LCase(S) Then
                grpNDX = K
                Exit Function
            End If
        Next


    End Function

    Public Function safeFilename(ByVal a) As String
        safeFilename = Replace(a, "\", "")
        safeFilename = Replace(safeFilename, "..", "")
    End Function


    Public Function listNDX(ByRef C As List(Of String), ByRef a$) As Integer
        Dim K As Integer = 0
        listNDX = 0

        For K = 0 To C.Count - 1
            If C(K).ToString = a Then
                listNDX = K + 1
                Exit Function
            End If
        Next

    End Function
    Public Function loadJSON(ByVal fileN$) As String
        loadJSON = ""
        If Dir(fileN) = "" Then Exit Function

        loadJSON = My.Computer.FileSystem.ReadAllText(fileN)

    End Function

    Public Function arrNDX(ByRef A$(), ByRef matcH$) As Integer
        'returns 0 if not found, otherwise NDX + 1
        Dim K As Long
        arrNDX = 0
        For K = 0 To UBound(A)
            If Trim(Str(A(K))) = matcH Then
                arrNDX = K + 1
                Exit Function
            End If
        Next
    End Function

    Public Function removeExtraSpaces(a) As String
        removeExtraSpaces = ""
        If Len(a) = 0 Then Exit Function
        Dim lastSpace As Boolean = False

        Dim K As Integer = 0
        For K = 0 To Len(a) - 1
            If lastSpace = False Then
                removeExtraSpaces += Mid(a, K + 1, 1)
            Else
                If Mid(a, K + 1, 1) <> " " Then removeExtraSpaces += Mid(a, K + 1, 1)
            End If
            If Mid(a, K + 1, 1) = " " Then
                lastSpace = True
            End If
        Next

    End Function

    Public Function countChars(a$, chr2Count$) As Integer
        countChars = 0

        Dim K As Integer = 0
        For K = 0 To Len(a) - 1
            If Mid(a, K + 1, 1) = chr2Count Then countChars += 1
        Next
    End Function

    Public Function stripToFilename(ByVal fileN$) As String
        'C:\Program Files\Checkmarx\Checkmarx Jobs Manager\Results\WebGoat.NET.Default 2014-10.9.2016-19.59.35.pdf
        stripToFilename = ""

        Do Until InStr(fileN, "\") = 0
            fileN = Mid(fileN, InStr(fileN, "\") + 1)
        Loop

        stripToFilename = fileN

    End Function

    Public Function addSlash(ByVal a$) As String
        addSlash = a
        If Len(a) = 0 Then Exit Function

        If Mid(a, Len(a), 1) <> "\" Then addSlash += "\"
    End Function

    Public Function getParentGroup(ByVal g$) As String
        Dim a$ = StrReverse(g)
        a = Mid(a, InStr(a, "\") + 1)
        Return StrReverse(a)
    End Function

    Public Function stripLastWord(ByVal g$) As String
        Dim a$ = StrReverse(g)
        a = Mid(a, InStr(a, " ") + 1)
        Return StrReverse(a)
    End Function

    Public Function getPDFappend(ByVal fileN$) As String
        'takes Cx filename used in calls to pull apart date section.. everything from -DD.M.YYYY-H.MM.SS.pdf- ---- cant assume '-' isnt used in filename
        'go to second - of reverse
        Dim a$ = StrReverse(fileN)
        Dim numDash As Integer = 0

        fileN = ""
        getPDFappend = ""

        If InStr(a, "-") = 0 Then Exit Function
        Dim K As Integer

        For K = 1 To Len(a)
            fileN += Mid(a, K, 1)
            If Mid(fileN, K, 1) = "-" Then numDash += 1
            If numDash = 2 Then Exit For
        Next

        getPDFappend = StrReverse(fileN)
    End Function

    Public Function assembleCollFromCLI(clI$) As Collection
        Dim C As New Collection
        ' takes windows dos-style dir output and makes sense of it for collection storage
        Dim tempStr$ = clI
        Dim K As Integer
        Do Until InStr(tempStr, "  ") = 0
            K = InStr(tempStr, "  ")
            If Len(Mid(tempStr, 1, K - 1)) Then C.Add(Mid(tempStr, 1, K - 1))
            tempStr = Replace(tempStr, Mid(tempStr, 1, K - 1) + "  ", "")
            'Debug.Print(tempStr)
        Loop
        tempStr = LTrim(tempStr)
        C.Add(Mid(tempStr, 1, InStr(tempStr, " ") - 1))
        tempStr = Replace(tempStr, Mid(tempStr, 1, InStr(tempStr, " ") - 1), "")
        C.Add(LTrim(tempStr))
        Return C

    End Function

    Public Function CSVtoCOLL(ByRef csV$) As Collection
        CSVtoCOLL = New Collection

        Dim splitCHR$ = ","
        If InStr(csV, splitCHR) = 0 Then splitCHR = ";"


        Dim longS = Split(csV, splitCHR)

        Dim K As Integer
        For K = 0 To UBound(longS)
            CSVtoCOLL.Add(longS(K))
        Next

    End Function


    Public Function CSVFiletoCOLL(ByRef csV$) As Collection
        CSVFiletoCOLL = New Collection
        If Dir(csV) = "" Then Exit Function

        'use file
        Dim FF As Integer
        FF = FreeFile()

        FileOpen(FF, csV, OpenMode.Input)

        Do Until EOF(FF) = True
            CSVFiletoCOLL.Add(LineInput(FF))
        Loop
        FileClose(FF)

    End Function

    Public Function argPROP(proP$) As String
        Dim a$ = ""
        argPROP = ""

        If Len(proP) = 0 Then Exit Function

        proP = LCase(proP)
        For Each arg In My.Application.CommandLineArgs
            a = LCase(arg)
            If InStr(a, "=") = 0 Then GoTo nextArg

            If proP = Mid(a, 1, InStr(a, "=") - 1) Then
                argPROP = Replace(a, proP + "=", "")
            End If
nextArg:
        Next
    End Function

    Public Sub safeKILL(ByRef fileN$)
        If Dir(fileN) <> "" Then Kill(fileN)
    End Sub


    Public Function filePROP(fileN$, proP$) As String
        filePROP = ""
        If Dir(fileN) = "" Then Exit Function

        If Len(proP) = 0 Then Exit Function

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""

        a = LineInput(FF)
        Do Until a = "" Or EOF(FF) = True
            If InStr(a, "=") = 0 Then GoTo nextLine

            If UCase(proP) = Mid(a, 1, InStr(a, "=") - 1) Then
                filePROP = Replace(a, proP + "=", "")
            End If
nextLine:
            a = LineInput(FF)
        Loop

        If Len(a) = 0 Then GoTo closeHere

        If UCase(proP) = Mid(a, 1, InStr(a, "=") - 1) Then
            filePROP = Replace(a, proP + "=", "")
        End If

closeHere:

        FileClose(FF)
    End Function

    Public Function allObjectsToList(fileN$) As List(Of String)
        allObjectsToList = New List(Of String)
        Dim C As New Collection
        Call getAllObjNamesFromFile(fileN, C)

        For Each A In C
            allObjectsToList.Add(loadOBJfromFILE(fileN, A))
        Next
    End Function

    Public Sub allObjectsWithProp(ByRef objS As List(Of String), prop$, propValue$, ByRef coll2Fill As Collection)
        coll2Fill = New Collection

        For Each O In objS
            If UCase(objProp(O, UCase(prop))) = UCase(propValue) Then coll2Fill.Add(objProp(O, "NAME"))
        Next
    End Sub

    Public Function typeOfParam(APIParamURI$, P As RestSharp.Parameter) As String
        typeOfParam = "QueryString"

        If InStr(APIParamURI$, "{" + P.Name + "}") Then
            typeOfParam = "Path"
        ElseIf InStr("$top,$count,$filter,$orderby,$select,$expand,-compare", P.Name) Then
            typeOfParam = "ODATA"
        End If

    End Function


    Public Sub getAllObjNamesFromFile(fileN$, ByRef collOFnames As Collection)
        collOFnames = New Collection

        If Dir(fileN) = "" Then Exit Sub

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""

        a = LineInput(FF)
        Do Until EOF(FF) = True
            If UCase(Mid(a, 1, 5)) = "NAME=" Then
                collOFnames.Add(Replace(a, "NAME" + "=", ""))
            End If
            a = LineInput(FF)
        Loop

        FileClose(FF)

    End Sub

    Public Function loadOBJfromFILE(fileN$, objName$) As String
        loadOBJfromFILE = ""

        If Dir(fileN) = "" Then Exit Function

        If Len(objName) = 0 Then Exit Function

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""
        Dim buildStr$ = ""

        Dim findSTR$ = "NAME=" + UCase(objName)

        a = LineInput(FF)
        Do Until UCase(a) = findSTR Or EOF(FF) = True
nextLine:
            a = LineInput(FF)
        Loop

        If UCase(a) = findSTR Then
            Do Until a = "" Or EOF(FF) = True
                buildStr += a + vbCrLf
                a = LineInput(FF)
            Loop
        End If

        loadOBJfromFILE = buildStr
        FileClose(FF)


    End Function

    Public Function objProp(ByRef ObjString As String, propName$) As String
        objProp = ""
        Dim findS$ = UCase(propName) + "="

        Dim O = Split(ObjString, vbCrLf)

        If UBound(O) = 0 Then Exit Function

        Dim K As Integer

        For K = 0 To UBound(O)
            If Mid(O(K), 1, Len(findS)) = UCase(propName) + "=" Then
                'found object, return property
                objProp = Mid(O(K), InStr(O(K), "=") + 1)
                Exit Function
            End If
        Next

    End Function

    Public Function xlsDataType(dType$) As String
        xlsDataType = "nonefound"
        Select Case dType
            Case "bigint", "int", "numeric", "float"
                xlsDataType = "Numeric"
            Case "datetime", "datetime2"
                xlsDataType = "DateTime"
            Case "date"
                xlsDataType = "Date"
            Case "time"
                xlsDataType = "Time"
            Case "bit"
                xlsDataType = "Boolean"
            Case "ntext", "nvarchar", "nchar", "varchar", "image", "uniqueidentifier", "real"
                xlsDataType = "String"
        End Select
        If xlsDataType = "nonefound" Then
            Debug.Print("No Def: " + dType)
            xlsDataType = "String"
        End If
    End Function

    Public Function xlsColName(colNum As Integer) As String
        Dim d As Integer
        Dim m As Integer
        Dim name As String
        d = colNum
        name = ""
        Do While (d > 0)
            m = (d - 1) Mod 26
            name = Chr(65 + m) + name
            d = Int((d - m) / 26)
        Loop
        xlsColName = name
    End Function


    Public NotInheritable Class Simple3Des
        Private TripleDes As TripleDESCryptoServiceProvider

        Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

            Dim sha1 As New SHA1CryptoServiceProvider

            ' Hash the key.
            Dim keyBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(key)
            Dim hash() As Byte = sha1.ComputeHash(keyBytes)

            ' Truncate or pad the hash.
            ReDim Preserve hash(length - 1)
            Return hash
        End Function

        Sub New(ByVal key As String)
            ' Initialize the crypto provider.
            TripleDes = New TripleDESCryptoServiceProvider
            TripleDes.Key = TruncateHash(key, TripleDES.KeySize \ 8)
            TripleDES.IV = TruncateHash("", TripleDES.BlockSize \ 8)
        End Sub

        Private Function EncryptData(ByVal plaintext As String) As String

            ' Convert the plaintext string to a byte array. 
            Dim plaintextBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream. 
            Dim encStream As New CryptoStream(ms,
                TripleDES.CreateEncryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string. 


            Return Convert.ToBase64String(ms.ToArray)

        End Function

        Private Function DecryptData(ByVal encryptedtext As String) As String

            ' Convert the encrypted text string to a byte array. 
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream. 
            Dim decStream As New CryptoStream(ms,
                TripleDES.CreateDecryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()

            ' Convert the plaintext stream to a string. 
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        End Function

        Public Function Decode(cipher As String) As String
            Try
                Return DecryptData(cipher)
            Catch ex As CryptographicException
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Function Encode(txt As String) As String
            Try
                Return EncryptData(txt)
            Catch ex As CryptographicException
                Throw New Exception(ex.Message)
            End Try
        End Function

    End Class
End Module
