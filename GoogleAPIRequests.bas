Attribute VB_Name = "Module2"
Option Explicit
Private p&, token, dic
Sub callAPI()

Dim token$, refresh$, authCode$, client_id$, scope$, secret$, apiUrl$, JSONString$
Dim apiCall As XMLHTTP60
Dim tokenJSON  As New Dictionary
Dim singleJSON As New Dictionary
Dim nestedJSON As New Dictionary
Dim requestJSON As New Collection
Dim dataRange As Range
Dim i As Long




'set credential values initially - set again after checking what's there
With ThisWorkbook.Sheets("secrets")
    authCode = .Range("B1").value
    token = .Range("B2").value
    refresh = .Range("B3").value
    scope = .Range("B4").value
    client_id = .Range("B5").value
    secret = .Range("B6").value
End With

'every time we call, need to check what access codes we have
If authCode = "" Then 'if no auth code then get one
    authCode = setAuthCode
    ThisWorkbook.Sheets("secrets").Range("B1").value = authCode
End If

If token = "" And refresh = "" Then 'if no token or refresh code then get both tokens
    Debug.Print "No Tokens"
    Set tokenJSON = setToken(authCode, secret, client_id)
    ThisWorkbook.Sheets("secrets").Range("B2").value = tokenJSON("obj.access_token")
    ThisWorkbook.Sheets("secrets").Range("B3").value = tokenJSON("obj.refresh_token")
    Set tokenJSON = Nothing
    
ElseIf Not refresh = "" Then 'if we have a refresh just get the access token
    Debug.Print "Refresh Token only"
    Set tokenJSON = setToken(authCode, secret, client_id, refresh)
    ThisWorkbook.Sheets("secrets").Range("B2").value = tokenJSON("obj.access_token")
    Set tokenJSON = Nothing
    
End If

With ThisWorkbook.Sheets("secrets")
    authCode = .Range("B1").value
    token = .Range("B2").value
    refresh = .Range("B3").value
    scope = .Range("B4").value
    client_id = .Range("B5").value
    secret = .Range("B6").value
End With

'calendar targeted with token in place
apiUrl = "https://www.googleapis.com/calendar/v3/calendars/INSERT CALENDAR REF HERE/events?access_token=" & token

'build JSON
Set dataRange = Sheets(1).cells(1, 1).CurrentRegion


For i = 2 To dataRange.Rows.Count
    With Sheets(1)

        singleJSON("summary") = .cells(i, 1)
        singleJSON("location") = "N/A"
        singleJSON("description") = .cells(i, 2)
        
        Set nestedJSON = New Dictionary
        nestedJSON("dateTime") = Format(.cells(i, 4), "yyyy-mm-dd") & "T" & .cells(i, 6) & "+00:00"
        nestedJSON("timeZone") = "Europe/London"
        singleJSON.Add "end", nestedJSON
        
        Set nestedJSON = New Dictionary
        nestedJSON("dateTime") = Format(.cells(i, 3), "yyyy-mm-dd") & "T" & .cells(i, 5) & "+00:00"
        nestedJSON("timeZone") = "Europe/London"
        singleJSON.Add "start", nestedJSON
        
        requestJSON.Add singleJSON
        Set singleJSON = Nothing
        
    End With
    
Next i

For i = 1 To requestJSON.Count
    Set apiCall = New XMLHTTP60
    'Debug.Print JSONString
    JSONString = JsonConverter.ConvertToJson(requestJSON(i), Whitespace:=3)

    'JSONString = Right(JSONString, Len(JSONString) - 1)
    'JSONString = Left(JSONString, Len(JSONString) - 1)
    
    apiCall.Open "POST", apiUrl, False
    apiCall.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    apiCall.send JSONString
    'JsonConverter.ConvertToJson(requestJSON, Whitespace:=3)
    Debug.Print apiCall.responseText

    
Next i


End Sub

'Get user to authorize access to their calendar, returns authorization code
Function setAuthCode() As String
Dim IE As InternetExplorer
Set IE = New InternetExplorer

Dim authUrl$, client_id$, scope$

MsgBox ("When prompted, navigate through the authorisation pages to allow the sheet to access your Google-held data. If successful, you will be presented with an authorisation code. Copy the code and paste it into the input box open on the sheet.")
With ThisWorkbook.Sheets("secrets")
    client_id = .Range("B5").value
    scope = .Range("B4").value
End With

authUrl = "https://accounts.google.com/o/oauth2/auth?response_type=code&client_id=" & client_id & "&redirect_uri=urn:ietf:wg:oauth:2.0:oob&scope=" & scope & "&access_type=offline"

With IE 'open browser to get user's permission
  .Navigate2 authUrl
End With

IE.Quit 'close the internet explorer instance

Dim authCode$
authCode = InputBox("Paste authorisation code here:")

setAuthCode = authCode

End Function

Function setToken(authCode As String, secret As String, client_id, Optional refresh) As Variant
Dim httpRequest As New MSXML2.XMLHTTP60
Dim tokenUrl$
Dim tokenArr() As Variant

If IsMissing(refresh) Then
    tokenUrl = "https://oauth2.googleapis.com/token?code=" & authCode & "&client_id=" & client_id & "&client_secret=" & secret & "&grant_type=authorization_code&redirect_uri=urn:ietf:wg:oauth:2.0:oob"
    httpRequest.Open "POST", tokenUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpRequest.send
    Set dic = ParseJson(httpRequest.responseText)
ElseIf Not refresh = "" Then
Debug.Print ("Refresh Called")
    tokenUrl = "https://oauth2.googleapis.com/token?refresh_token=" & refresh & "&client_id=" & client_id & "&client_secret=" & secret & "&redirect_uri=urn:ietf:wg:oauth:2.0:oob&grant_type=refresh_token"
    httpRequest.Open "POST", tokenUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpRequest.send
    Set dic = ParseJson(httpRequest.responseText)
Else
'handle error
End If


ReDim tokenArr(dic.Count)

Set setToken = dic

End Function




'-------------------------------------------------------------------
' VBA JSON Parser
'-------------------------------------------------------------------

Function ParseJson(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJson = dic
    Set dic = Nothing
End Function
Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function
Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function
'-------------------------------------------------------------------
' Support Functions
'-------------------------------------------------------------------
Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function
Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function
Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function
Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function
Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    Debug.Print s
End Function
Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.Keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.Keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function
Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function

