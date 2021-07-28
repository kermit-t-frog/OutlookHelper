Attribute VB_Name = "CheckCalendar"
Option Explicit

Sub OverlapMenu()
Ueberlapp.Show
End Sub

Public Function Overlap(recipients As Object, _
                        ByVal FirstDate As Date, _
                        ByVal durationMinutes As Integer, _
                        Optional ByVal startOfDay As Date = #8:30:00 AM#, _
                        Optional ByVal endOfDay As Date = #4:30:00 PM#, _
                        Optional ByVal lunchStart As Date = #12:00:00 PM#, _
                        Optional ByVal lunchEnd As Date = #1:00:00 PM#, _
                        Optional ByVal endOfFriday As Date = #3:00:00 PM#, _
                        Optional tentativeElesewhereIsAvailable As Boolean = True, _
                        Optional ByVal MaxDistanceDays As Integer = 5, _
                        Optional ByVal ResolutionInMinutes As Integer = 15)
    ' set the defaults
    'If startOfDay = 0 Then startOfDay = TimeSerial(8, 30, 0)
    'If endOfDay = 0 Then endOfDay = TimeSerial(16, 30, 0)
    'If endOfFriday = 0 Then endOfFriday = TimeSerial(15, 15, 0)
    'If lunchStart = 0 Then lunchStart = TimeSerial(12, 0, 0)
    'If lunchEnd = 0 Then lunchEnd = TimeSerial(13, 0, 0)
    
    Dim LastDate As Date:           LastDate = DateAdd("D", FirstDate, MaxDistanceDays)
    Dim stepsPerDay As Integer:     stepsPerDay = 24 * 60 / ResolutionInMinutes
    Dim N As Integer:               N = (LastDate - FirstDate + 1) * stepsPerDay
    
    ' https://docs.microsoft.com/en-us/office/vba/api/outlook.olbusystatus
    ' Enum OlBusyStatus
    ' olFree             0 user is available
    ' olTentative        1 user has a tentative appointment scheduled
    ' olBusys            2 user is busy
    ' olOutOfOffice      3 user is out of office
    ' olWorkingElsewhere 4 user is working in a location away from office

    
    Dim availabilities As String, availability As String, bracket As Date
    Dim i As Long
    ' prepare the availability string based on calendar, working times, lunch times etc
    
    availabilities = Replace(Space(N), " ", "0")
    For i = 1 To N
        ' bracket is LEFT end of time slot
        bracket = DateAdd("n", ((i - 1) * ResolutionInMinutes), FirstDate)
        If Weekday(bracket, vbMonday) > 5 Or TimeValue(bracket) < startOfDay Or TimeValue(bracket) >= endOfDay _
        Or (TimeValue(bracket) >= lunchStart And TimeValue(bracket) < lunchEnd) _
        Or (Weekday(bracket, vbMonday) = 5 And TimeValue(bracket) >= endOfFriday) _
        Or bracket < Now Then
            Mid(availabilities, i, 1) = "2"
        End If
    Next i
    
    Dim slotAll As String, slotI As String
    Dim rec As Outlook.Recipient
    
    For Each rec In recipients ' itm.recipients
        On Error GoTo nextRec ' FreeBusyMethod not always available
        availability = Left(rec.FreeBusy(FirstDate, ResolutionInMinutes, True), N)  ' next x days, hourly resolution
        For i = 1 To N
            slotAll = Mid(availabilities, i, 1) ' i'th state of the group
            slotI = Mid(availability, i, 1)     ' i'th state of the new resource
            Mid(availabilities, i, 1) = JointAvailability(slotAll, slotI, tentativeElesewhereIsAvailable)
        Next i
nextRec:
    On Error GoTo 0
    Next rec
    
    Dim out As Object: Set out = Outlook.Application.CreateObject("Scripting.Dictionary")
    Dim minimumLength As Integer:   minimumLength = durationMinutes / ResolutionInMinutes
    Dim runStart As Long, runLength As Long
    For i = 1 To N
        If Mid(availabilities, i, 1) = "0" Then
            If runStart = 0 Then runStart = i
            runLength = runLength + 1
        Else
            If runStart <> 0 Then
                If runLength >= minimumLength Then
                    out.Add DateAdd("n", (runStart - 1) * ResolutionInMinutes, FirstDate), DateAdd("n", ((runStart - 1 + runLength)) * ResolutionInMinutes, FirstDate)
                End If
                runStart = 0
                runLength = 0
            End If
        End If
    Next i

    Set Overlap = out: Set out = Nothing
End Function

' returns a joint business flag in string fromat from OlBusyStatus
' https://docs.microsoft.com/en-us/office/vba/api/outlook.olbusystatus
' Enum OlBusyStatus
' olFree             0 user is available
' olTentative        1 user has a tentative appointment scheduled
' olBusys            2 user is busy
' olOutOfOffice      3 user is out of office
' olWorkingElsewhere 4 user is working in a location away from office
' cases
' case                  result          count   total
' ---------------------------------------------------
' A = B                 A               5       5
' A = 0                 B               4       9
' B = 0                 A               4       13
' A = 2, A = 3          A               6       19
' B = 2, B = 3          B               4       23
' A=(1,4) or B=(4,1)    1               2       25 / 25
Private Function JointAvailability(resourceA As String, resourceB As String, Optional tentativeElesewhereIsAvailable As Boolean = False) As String
    Dim out As String
    Select Case True
        Case resourceA = resourceB, resourceA = "2", resourceA = "3", resourceB = "0"
                out = resourceA
            Case resourceA = "0", resourceB = "2", resourceB = "3"
                out = resourceB
            Case resourceB = "1", resourceA = "1"
                out = "1"
            Case Else
                Err.Raise 999, , "unknown combination"
        End Select
    If tentativeElesewhereIsAvailable And (out = "1" Or out = "4") Then out = "0"
    JointAvailability = out
End Function
