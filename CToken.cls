VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing the MS Graph access token for the MS Teams app.
'
Option Explicit

Private access_token_ As String
Private refresh_token_ As String

Private access_token_location_ As String
Private refresh_token_location_ As String

Private tenantid_ As String
Private clientid_ As String
Private permissionscope_ As String
Private redirecturi_ As String

Private username_ As String
Private password_ As String 
Private clientsecret_ As String 

Private apistate_ As String
Private access_token_lifetime_minutes As Integer '  according to graph spec: 3599 seconds
Private refresh_token_lifetime_days_ As Integer  '  according to graph spec: 13   days 
'

Public Sub init(tenantid As String, _
                clientid As String, _
                client_secret As String, _
                scope As String, _
                redirecturi As String, _
                access_token_location As String, _
                refresh_token_location As String, _
                username As String, _
                Optional ByVal tokenlifetime_minutes As Integer = 59, _
                Optional ByVal refreshtokenlifetime_days As Integer = 13, _
                Optional ByVal apistate As String = "12345")
    
    tenantid_ = tenantid
    clientid_ = clientid
    access_token_location_ = access_token_location
    refresh_token_location_ = refresh_token_location
    permissionscope_ = scope
    redirecturi_ = redirecturi
    apistate_ = apistate
    access_token_lifetime_minutes = tokenlifetime_minutes
    refresh_token_lifetime_days_ = refreshtokenlifetime_days
    clientsecret_ = client_secret
    username_ = username
    
    If Dir(access_token_location_) <> "" Then
        Open access_token_location_ For Input As #1
            Line Input #1, access_token_
        Close #1
    End If
    If Dir(refresh_token_location_) <> "" Then
        Open refresh_token_location_ For Input As #1
            Line Input #1, refresh_token_
        Close #1
    End If
End Sub

Public Property Get access_token() As String
    If IsInvalidAccessToken Then
        Call NewToken
    Else
        Open access_token_location_ For Input As #1
            Line Input #1, access_token_
        Close #1
    End If
        access_token = access_token_
End Property

Private Property Get IsInvalidAccessToken() As Boolean
    IsInvalidAccessToken = True
    If Dir(access_token_location_) <> "" Then
        If DateDiff("s", Now(), FileDateTime(access_token_location_) + TimeSerial(0, access_token_lifetime_minutes, 0)) > 0 Then
            IsInvalidAccessToken = False
        End If
    End If
End Property

Private Property Get IsInvalidRefreshToken() As Boolean
    IsInvalidRefreshToken = True
    If Dir(access_token_location_) <> "" Then
        If DateDiff("D", Date, FileDateTime(refresh_token_location_) + refresh_token_lifetime_days_) > 0 Then
            IsInvalidRefreshToken = False
        End If
    End If
End Property

Private Sub StoreAccessToken(t As String)
    If Dir(access_token_location_) <> "" Then Kill access_token_location_
    Open access_token_location_ For Output As #1
        Print #1, t
    Close #1
End Sub

Private Sub StoreRefreshToken(t As String)
    If Dir(refresh_token_location_) <> "" Then Kill refresh_token_location_
    Open refresh_token_location_ For Output As #1
        Print #1, t
    Close #1
End Sub

Private Sub NewToken()
    Dim ht As Object, url As String, load As String
    url = "https://login.microsoftonline.com/" & tenantid_ & "/oauth2/v2.0/token"
    
    If IsInvalidRefreshToken Then
        If password_ = "" Then
            Dim pw As New PWBox         ' das ist noch unsauber.
            pw.Show                     '
            password_ = pw.Value        '
            Unload pw: Set pw = Nothing '
            If password_ = "" Then Err.Raise 999, "", "No password provided."
        End If
        
        load = "client_id=" & clientid_ _
             & "&scope=" & permissionscope_ _
             & "&grant_type=password" _
             & "&client_secret=" & clientsecret_ _
             & "&username=" & username_ _
             & "&password=" & password_
    Else
        load = "client_id=" & clientid_ _
             & "&scope=" & permissionscope_ _
             & "&grant_type=refresh_token" _
             & "&client_secret=" & clientsecret_ _
             & "&refresh_token=" & refresh_token_ _
             & "&redirect_uri=" & redirecturi_
    End If
        
    Set ht = Outlook.Application.CreateObject("WinHttp.WinHttpRequest.5.1")
    ht.SetProxy 2, "proxy.d3622070.lfilsa.nordlb.local:8080"
    ht.Open "POST", url, False
    ht.SetRequestHeader "Host", "login.microsoftonline.com"
    ht.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    ht.Option(6) = False 'REDIRECT => FALSE, wird ggf. nicht mehr benötigt.
    ht.Send load
    ht.WaitForResponse
    
    If InStr(1, ht.ResponseText, "access_token") Then
        access_token_ = Left(Mid(ht.ResponseText, InStr(1, ht.ResponseText, "access_token") + 15), 3000)
        access_token_ = Left(access_token_, InStr(1, access_token_, """") - 1)
    Else
        Err.Raise 999, "", "Cannot obtain access_token."
    End If

    If InStr(1, ht.ResponseText, "refresh_token") Then
        refresh_token_ = Left(Mid(ht.ResponseText, InStr(1, ht.ResponseText, "refresh_token") + 16), 3000)
        refresh_token_ = Left(refresh_token_, InStr(1, refresh_token_, """") - 1)
    Else
        Err.Raise 999, "", "Cannot obtain refresh_token."
    End If
    Set ht = Nothing
    Call StoreAccessToken(access_token_)
    Call StoreRefreshToken(refresh_token_)
    
End Sub

