VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ueberlapp 
   Caption         =   "Überlapp"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   OleObjectBlob   =   "Ueberlapp.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Ueberlapp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private recipients_ As Outlook.recipients
Private start_of_day_ As Date
Private end_of_day_ As Date
Private end_of_friday_ As Date
Private lunch_start_ As Date
Private lunch_end_ As Date
Private result_ As Object

Private Sub btn_AddRecipients_Click()
    Dim snd As Outlook.SelectNamesDialog
    Dim al As AddressList
    Set snd = Application.Session.GetSelectNamesDialog
    Set al = Outlook.Application.GetNamespace("MAPI").AddressLists("Global Address List")

    With snd
        .InitialAddressList = al
        
        .recipients = recipients_
        .SetDefaultDisplayMode olDefaultMeeting
        .ShowOnlyInitialAddressList = True
        .Display
    End With
    'Dim rec As Recipient
    'For Each rec In snd.recipients
     '   recipients_.Add rec
    'Next rec
    Set recipients_ = snd.recipients
    Set al = Nothing: Set snd = Nothing
    
    Call RefreshRecipients

End Sub

Private Sub btn_GetOverlap_Click()

If recipients_.Count = 0 Then
    MsgBox "No participants.", vbInformation + vbOKOnly
    GoTo exitSub
End If
Dim output As Object

If Not Me.chk_lunch Then
    lunch_start_ = #10:00:00 PM#
    lunch_end_ = #10:01:00 PM#
End If

Set output = KalenderPruefung.Overlap(recipients_, _
            Me.lst_Startdatum.SelStart, _
            GetSelectedDuration, _
            start_of_day_, _
            end_of_day_, _
            lunch_start_, _
            lunch_end_, _
            end_of_friday_, _
            Me.chk_AwayIsAvailable.Value, _
            DateDiff("D", Me.lst_Startdatum.SelStart, Me.lst_Startdatum.SelEnd))
Me.lst_results.Clear

Dim slot As Variant
Set result_ = Outlook.Application.CreateObject("Scripting.Dictionary")
For Each slot In output
    Me.lst_results.AddItem Format(slot, "ddd, YY-MM-DD hh:mm") & " - " & Format(output(slot), "hh:mm")
    result_.Add Format(slot, "ddd, YY-MM-DD hh:mm") & " - " & Format(output(slot), "hh:mm"), slot
Next slot
exitSub:
End Sub

Private Sub btn_remove_recipients_Click()

Dim rec As Recipient
For Each rec In recipients_
If rec.Name = Me.lst_Recipients.Value Then rec.Delete
Next rec
Call RefreshRecipients

End Sub

Private Sub RefreshRecipients()
Me.lst_Recipients.Clear
  Dim rec As Outlook.Recipient
    For Each rec In recipients_
        Me.lst_Recipients.AddItem rec.Name
    Next rec
    Me.Repaint
    Set rec = Nothing
End Sub

Private Function GetSelectedDuration() As Integer
    Select Case True
        Case Me.opt_duration_15
            GetSelectedDuration = 15
        Case Me.opt_duration_30
            GetSelectedDuration = 30
        Case Me.opt_duration_45
            GetSelectedDuration = 45
        Case Me.opt_duration_60
            GetSelectedDuration = 60
        Case Me.opt_duration_90
            GetSelectedDuration = 90
        Case Else
            Err.Raise 999, , "not implemented"
    End Select
End Function


Private Sub chk_lunch_Click()
Me.txt_LunchStart.Enabled = Not Me.txt_LunchStart.Enabled
Me.txt_LunchEnd.Enabled = Not Me.txt_LunchEnd.Enabled
End Sub



Private Sub lst_results_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim apo As Outlook.AppointmentItem

Set apo = Outlook.Application.CreateItem(olAppointmentItem)
apo.MeetingStatus = olMeeting
apo.start = result_(Me.lst_results.Value)
apo.Duration = GetSelectedDuration

Dim rec As Recipient
For Each rec In recipients_
    If rec.Address <> Outlook.Application.GetNamespace("MAPI").CurrentUser.Address Then
        apo.recipients.Add rec.AddressEntry
    End If
Next rec
apo.Display
Unload Me
End Sub

Private Sub txt_EndOfDay_AfterUpdate()
    If Not IsDate(Me.txt_EndOfDay.Value) Then
        MsgBox "Keine zulässige Uhrzeit.", vbExclamation + vbOKOnly
        Me.txt_EndOfDay.Value = end_of_day_
    Else
        end_of_day_ = CDate(Me.txt_EndOfDay.Value)
    End If
End Sub

Private Sub txt_EndOfDay_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txt_EndOfDay.Value) Then
        end_of_day_ = CDate(Me.txt_EndOfDay.Value)
    End If
End Sub

Private Sub txt_StartOfDay_AfterUpdate()
    If Not IsDate(Me.txt_StartOfDay.Value) Then
        MsgBox "Keine zulässige Uhrzeit.", vbExclamation + vbOKOnly
        Me.txt_StartOfDay.Value = start_of_day_
    Else
        start_of_day_ = CDate(Me.txt_StartOfDay.Value)
    End If
End Sub

Private Sub txt_StartOfDay_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txt_StartOfDay.Value) Then
        start_of_day_ = CDate(Me.txt_StartOfDay.Value)
    End If
End Sub

Private Sub txt_EndOfFriday_AfterUpdate()
    If Not IsDate(Me.txt_EndOfFriday.Value) Then
        MsgBox "Keine zulässige Uhrzeit.", vbExclamation + vbOKOnly
        Me.txt_EndOfFriday.Value = end_of_friday_
    Else
        end_of_friday_ = CDate(Me.txt_EndOfFriday.Value)
    End If
End Sub

Private Sub txt_EndOfFriday_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txt_EndOfFriday.Value) Then
        end_of_friday_ = CDate(Me.txt_EndOfFriday.Value)
    End If
End Sub

Private Sub txt_LunchStart_AfterUpdate()
    If Not IsDate(Me.txt_LunchStart.Value) Then
        MsgBox "Keine zulässige Uhrzeit.", vbExclamation + vbOKOnly
        Me.txt_LunchStart.Value = lunch_start_
    Else
        lunch_start_ = CDate(Me.txt_LunchStart.Value)
    End If
End Sub

Private Sub txt_LunchStart_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txt_LunchStart.Value) Then
        lunch_start_ = CDate(Me.txt_LunchStart.Value)
    End If
End Sub

Private Sub txt_LunchEnd_AfterUpdate()
    If Not IsDate(Me.txt_LunchEnd.Value) Then
        MsgBox "Keine zulässige Uhrzeit.", vbExclamation + vbOKOnly
        Me.txt_LunchEnd.Value = lunch_end_
    Else
        lunch_end_ = CDate(Me.txt_LunchEnd.Value)
    End If
End Sub

Private Sub txt_LunchEnd_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txt_LunchEnd.Value) Then
        lunch_end_ = CDate(Me.txt_LunchEnd.Value)
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim apo As Outlook.AppointmentItem
    Set apo = Outlook.Application.CreateItem(olAppointmentItem)
    Set recipients_ = apo.recipients
    recipients_.Add Outlook.Application.GetNamespace("MAPI").CurrentUser.Name
    Set apo = Nothing
    start_of_day_ = #8:30:00 AM#
     end_of_day_ = #4:30:00 PM#
    end_of_friday_ = #3:00:00 PM#
    lunch_start_ = #12:00:00 PM#
    lunch_end_ = #1:00:00 PM#
    Me.txt_StartOfDay.Value = start_of_day_
    Me.txt_EndOfDay.Value = end_of_day_
    Me.txt_EndOfFriday.Value = end_of_friday_
    Me.txt_LunchStart.Value = lunch_start_
    Me.txt_LunchEnd.Value = lunch_end_
    Call RefreshRecipients
End Sub

