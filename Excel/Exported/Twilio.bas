Attribute VB_Name = "Twilio"
Option Explicit

Const NOINTERNETAVAILABLE = -2147012889

' **************************************************
' NOTE: You will need to update these constants
'  with your own Twilio credentials
Const ACCOUNTSID As String = "[YOUR_ACCOUNT_SID]"
Const AUTHTOKEN As String = "[YOUR_AUTH_TOKEN]"
' **************************************************

Const BASEURL As String = "https://api.twilio.com"

Public Sub conSendSms(control As IRibbonControl)
    With frmSendSms
        .Show
    End With
End Sub

Function SendMessage(fromNumber As String, toNumber As String, body As String)
    Dim MessageUrl As String
    
    On Error GoTo Error_Handler
    
    ' setup the URL
    MessageUrl = BASEURL & "/2010-04-01/Accounts/" & ACCOUNTSID & "/Messages"
    
    ' setup the request and authorization
    Dim http As MSXML2.XMLHTTP60
    Set http = New MSXML2.XMLHTTP60
    
    http.Open "POST", MessageUrl, False, ACCOUNTSID, AUTHTOKEN
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    Dim postData As String
    postData = "From=" & URLEncode(fromNumber) _
                & "&To=" & URLEncode(toNumber) _
                & "&Body=" & body
    
    Debug.Print postData
    
    ' send the POST data
    http.send postData
    
    ' optionally write out the response if you need to check if it worked
    Debug.Print http.responseText
    
    If http.Status = 201 Then

    ElseIf http.Status = 400 Then
        MsgBox "Failed with error# " & _
            http.Status & _
            " " & http.statusText & vbCrLf & vbCrLf & _
            http.responseText
    ElseIf http.Status = 401 Then
        MsgBox "Failed with error# " & http.Status & _
            " " & http.statusText & vbCrLf & vbCrLf
    Else
        MsgBox "Failed with error# " & http.Status & _
            " " & http.statusText
    End If

Exit_Procedure:

    On Error Resume Next

    ' clean up
    Set http = Nothing

Exit Function

Error_Handler:

    Select Case Err.Number

        Case NOINTERNETAVAILABLE
            MsgBox "Connection to the internet cannot be made or " & _
                "Twilio website address is wrong"

        Case Else
            MsgBox "Error: " & Err.Number & "; Description: " & Err.Description

            Resume Exit_Procedure

        Resume

    End Select
End Function

Function MakeCall(fromNumber As String, toNumber As String, message As String)
    Dim CallUrl As String
    
    On Error GoTo Error_Handler
    
    ' setup the URL
    CallUrl = BASEURL & "/2010-04-01/Accounts/" & ACCOUNTSID & "/Calls"
    
    ' setup the request and authorization
    Dim http As MSXML2.XMLHTTP60
    Set http = New MSXML2.XMLHTTP60
    
    http.Open "POST", CallUrl, False, ACCOUNTSID, AUTHTOKEN
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    Dim postData As String
    postData = "From=" & URLEncode(fromNumber) _
                & "&To=" & URLEncode(toNumber) _
                & "&Url=http://twimlets.com/message?Message=" & URLEncode(URLEncode(message, True))
    
    Debug.Print postData
    
    ' send the POST data
    http.send postData
    
    ' optionally write out the response if you need to check if it worked
    Debug.Print http.responseText
    
    If http.Status = 201 Then

    ElseIf http.Status = 400 Then
        MsgBox "Failed with error# " & _
            http.Status & _
            " " & http.statusText & vbCrLf & vbCrLf & _
            http.responseText
    ElseIf http.Status = 401 Then
        MsgBox "Failed with error# " & http.Status & _
            " " & http.statusText & vbCrLf & vbCrLf
    Else
        MsgBox "Failed with error# " & http.Status & _
            " " & http.statusText
    End If

Exit_Procedure:

    On Error Resume Next

    ' clean up
    Set http = Nothing

Exit Function

Error_Handler:

    Select Case Err.Number

        Case NOINTERNETAVAILABLE
            MsgBox "Connection to the internet cannot be made or " & _
                "Twilio website address is wrong"

        Case Else
            MsgBox "Error: " & Err.Number & "; Description: " & Err.Description

            Resume Exit_Procedure

        Resume

    End Select
End Function

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

