Attribute VB_Name = "Teams"
' Create a Teams Calendar entry via automation. Useful if there is no teams integration available in Outlook.
' Preparation:
'
' 1. Go to https://portal.azure.com/ -> "App registration".
' App registration  :   Name the app
'                       choose "single tenant"
'                       redirect to "http://localhost"
' Secrets           :   We need a client secret
' API permissions   :   Calendars.ReadWrite, offline_access
' Redirect URI      :   "http://localhost"
' take note of      :   Directory (tenant) ID ("tenant")
'                       Application (client) ID ("client")
'                       Client secret ("secret")
'
' Description       :   After successfully negoatiating the OAUTH2.0 authentification, this
'                       procedure lets you set a Teams Online Meeting at the date/time of your liking
'
' Preliminaries     :   After setting your app in azure, browse to the following page and log-in. This
'                       is only required once. 'https://login.microsoftonline.com/{TENANTID}/oauth2/v2.0/authorize?client_id={CLIENTID}&response_type=code&redirect_uri=http%3A%2F%2Flocalhost&response_mode=query&scope=offline_access%20Calendars.ReadWrite&state=12345
'
' Version 2021-07-27
'

Option Explicit
Private Const CONF_FILE_LOCATION As String = "C:\path\to\teams.conf"
'


Public Sub CreateEventFromAppointmentItem()
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Do you want to create a Teams Online Meeting?", vbYesNo + vbQuestion)
    If answer = vbNo Then GoTo exitSub
    
    Dim itm As Outlook.AppointmentItem, recip As Outlook.Recipient, pa As Outlook.PropertyAccessor
    Set itm = Outlook.Application.ActiveInspector.CurrentItem
    
    Dim conf As Object
    Set conf = ReadConfig(CONF_FILE_LOCATION)
    
    Dim organizer As String, startTime As Date, endTime As Date, subject As String
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    
    Dim token As New CToken
    token.init tenantid:=conf("tenant_id"), _
               clientid:=conf("client_id"), _
               client_secret:=conf("client_secret"), _
               scope:=conf("permission_scope"), _
               redirecturi:=conf("redirect_uri"), _
               access_token_location:=conf("access_token_location"), _
               refresh_token_location:=conf("refresh_token_location"), _
               username:=conf("username")
    
    Set recip = itm.recipients(1)
    Set pa = recip.PropertyAccessor
    organizer = pa.GetProperty(PR_SMTP_ADDRESS)
    startTime = itm.start
    endTime = itm.End
    subject = itm.subject
    
    Dim response As String
    response = CreateEvent(organizer, subject, startTime, endTime, token.access_token)
    
    itm.Body = itm.Body & "An Konferenz teilnehmen:" & ExtractValue(response, "joinurl")
    itm.location = ExtractValue(response, "tollnumber") & ",," & ExtractValue(response, "conferenceid") & "#"

    Set recip = Nothing: Set pa = Nothing:  Set itm = Nothing: Set token = Nothing: Set conf = Nothing
exitSub:
End Sub

Private Function CreateEvent(organizer_address As String, subject As String, startTime As Date, endTime As Date, token As String) As String
    Dim attendees As String, jsonbody As String
    attendees = "{""emailAddress"":{""address"":""" & organizer_address & """,""name"":""" & organizer_address & """},""type"":""required""}"
    jsonbody = "{""subject"":""" & subject & """," _
                & """start"":{""datetime"":""" & Format(startTime, "YYYY-MM-DDTHH:mm:ss") & """,""timeZone"":""W. Europe Standard Time""}," _
                & """end"":{""datetime"":""" & Format(endTime, "YYYY-MM-DDTHH:mm:ss") & """,""timeZone"":""W. Europe Standard Time""}," _
                & """attendees"":[" & attendees & "]," _
                & """isOnlineMeeting"":true," _
                & """onlineMeetingProvider"":""teamsForBusiness""}"
    
    Dim ht As Object: Set ht = Outlook.Application.CreateObject("MSXML2.XMLHTTP") ' kurzer Dienstweg.
    With ht
        .Open "POST", "https://graph.microsoft.com/v1.0/me/events", False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & token
        .Send jsonbody
        CreateEvent = .ResponseText
    End With: Set ht = Nothing
End Function

Private Function ExtractValue(fromString As String, key As String, Optional closingToken As String = """") As String
    If InStr(1, fromString, key, vbTextCompare) = 0 Then Err.Raise 999, , "Cannot find key in string"
    Dim subs As String
    subs = Mid(fromString, InStr(1, fromString, key, vbTextCompare) + Len(key) + 3, 999)
    subs = Left(subs, InStr(1, subs, closingToken) - 1)
    ExtractValue = subs
End Function

' Config file is expected to be formated as "key:value". Key may not contain a colon, key+value must be on the same line.
Private Function ReadConfig(file As String) As Object ' Returns a Scripting.Dictionary
    Dim dct As Object, FileNum As Integer, dataline As String
    Set dct = Outlook.Application.CreateObject("Scripting.Dictionary")
    FileNum = FreeFile
    Open file For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, dataline
        If InStr(1, dataline, ":") > 0 Then
            dct.Add Left(dataline, InStr(1, dataline, ":") - 1), Mid(dataline, InStr(1, dataline, ":") + 1, 99999)
        End If
    Loop
    Close #FileNum
    Set ReadConfig = dct: Set dct = Nothing
End Function
