Function slaCalculate(receivedTimeString As String, Project As String, Task As String, subTask As String)
Dim slaTimeLimit As Long
Dim addType As String
Dim supportDays As Integer
Dim ShiftStartTimeString As String
Dim ShiftEndTimeString As String
Dim receivedTime, Workinghrs, ShiftStartTime, ShiftEndTime, FridayDate As Date
Dim tempHrCal, tempHrCal2, WeekendHrs, SettingLastRow As Long

SettingLastRow = ThisWorkbook.Sheets("Setting").Cells(ThisWorkbook.Sheets("Setting").Rows.Count, "B").End(xlUp).Row

For I = 1 To SettingLastRow
    If ThisWorkbook.Sheets("Setting").Range("B" & I).Value = Project And ThisWorkbook.Sheets("Setting").Range("C" & I).Value = subTask And ThisWorkbook.Sheets("Setting").Range("D" & I).Value = Task Then
        addType = ThisWorkbook.Sheets("Setting").Range("F" & I).Value
        supportDays = ThisWorkbook.Sheets("Setting").Range("G" & I).Value
        slaTimeLimit = ThisWorkbook.Sheets("Setting").Range("E" & I).Value
        ShiftStartTimeString = ThisWorkbook.Sheets("Setting").Range("H" & I).Value
        ShiftEndTimeString = ThisWorkbook.Sheets("Setting").Range("I" & I).Value
        Exit For
    End If
Next

receivedTime = CDate(receivedTimeString)
ShiftStartTime = CDate(ShiftStartTimeString)
ShiftEndTime = CDate(ShiftEndTimeString)

'If received time is in non-working hrs make it to start of next working hrs or monday morning if received in weekend
If TimeSerial(Hour:=Hour(receivedTime), Minute:=Minute(receivedTime), Second:=Second(receivedTime)) > TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime)) Then
    If Weekday(receivedTime, vbFriday) = 1 Then
        receivedTime = DateAdd("d", 3, receivedTime)
    ElseIf Weekday(receivedTime, vbSaturday) = 1 Then
        receivedTime = DateAdd("d", 2, receivedTime)
    Else
        receivedTime = DateAdd("d", 1, receivedTime)
    End If
    receivedTime = CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftStartTime), Minute:=Minute(ShiftStartTime), Second:=Second(ShiftStartTime)))
ElseIf TimeSerial(Hour:=Hour(receivedTime), Minute:=Minute(receivedTime), Second:=Second(receivedTime)) < TimeSerial(Hour:=Hour(ShiftStartTime), Minute:=Minute(ShiftStartTime), Second:=Second(ShiftStartTime)) Then
    If Weekday(receivedTime, vbSaturday) = 1 Then
        receivedTime = DateAdd("d", 2, receivedTime)
    ElseIf Weekday(receivedTime, vbSunday) = 1 Then
        receivedTime = DateAdd("d", 1, receivedTime)
    End If
    receivedTime = CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftStartTime), Minute:=Minute(ShiftStartTime), Second:=Second(ShiftStartTime)))
End If

'convert sla to minutes for easy calculation
If addType = "hours" Then slaTimeLimit = slaTimeLimit * 60
If addType = "days" Then slaTimeLimit = slaTimeLimit * (24 * 60)
'Calculating hours to add with non-working hours
tempHrCal = DateDiff("h", receivedTime, CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime))))
tempHrCal = (slaTimeLimit / 60) - tempHrCal
tempHrCal2 = WorksheetFunction.RoundUp(tempHrCal / DateDiff("h", ShiftStartTime, ShiftEndTime), 0)
tempHrCal2 = tempHrCal2 * (24 - (DateDiff("h", ShiftStartTime, ShiftEndTime)))
tempHrCal2 = (tempHrCal2 * 60) + slaTimeLimit

'For calculating weekends
If supportDays = 7 Then
    WeekendHrs = 0
ElseIf supportDays = 5 Then
    FridayDate = DateAdd("d", 8 - Weekday(receivedTime, vbFriday), receivedTime)
    FridayDate = CDate(DateSerial(Year:=Year(FridayDate), Month:=Month(FridayDate), Day:=Day(FridayDate)) & " " & TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime)))
    
   ' WeekendHrs = (TimeSerial(24, 0, 0)) * (DateDiff("h", ShiftStartTime, ShiftEndTime)) - WorksheetFunction.NetworkDays(receivedTime, FridayDate)
    WeekendHrs = DateDiff("h", ShiftStartTime, ShiftEndTime) * DateDiff("d", receivedTime, FridayDate)

    If (WorksheetFunction.NetworkDays(receivedTime, FridayDate) = 1) Then
        If UBound(Split(CStr(CDbl(receivedTime)), ".")) <= 0 Then
            WeekendHrs = WeekendHrs + WorksheetFunction.Median(0, ShiftEndTime, ShiftStartTime)
        Else
            WeekendHrs = WeekendHrs + WorksheetFunction.Median(Split(CStr(CDbl(FridayDate)), ".")(1), ShiftStartTime, ShiftEndTime)
        End If
    Else
        WeekendHrs = WeekendHrs + (DateDiff("h", ShiftStartTime, ShiftEndTime))
    End If
    
    If UBound(Split(CStr(CDbl(receivedTime)), ".")) <= 0 Then
        WeekendHrs = WeekendHrs - WorksheetFunction.Median(WorksheetFunction.NetworkDays(receivedTime, receivedTime) * 0, ShiftEndTime, ShiftStartTime)
    Else
        WeekendHrs = WeekendHrs - WorksheetFunction.Median(WorksheetFunction.NetworkDays(receivedTime, receivedTime) * Split(CStr(CDbl(receivedTime)), ".")(1), ShiftEndTime, ShiftStartTime)
    End If
    
    WeekendHrs = WorksheetFunction.RoundUp(((slaTimeLimit / 60) - WeekendHrs) / (DateDiff("h", ShiftStartTime, ShiftEndTime) * 5), 0)
    If WeekendHrs < 0 Then WeekendHrs = 0
    WeekendHrs = WeekendHrs * 48
End If

'Answer = receivedtime + (total hours including non working hrs) + weekends if any
slaCalculate = receivedTime + TimeSerial(0, tempHrCal2, 0) + TimeSerial(WeekendHrs, 0, 0)

 '-------------------Only use For Calculating total hrs from 'Start' date to 'End' date--------------------------------

'
'Result = (WorksheetFunction.NetworkDays(receivedTime, endTime) - 1) * (ShiftEndTime - ShiftStartTime)
'
'If (WorksheetFunction.NetworkDays(endTime, endTime) = 1) Then
'    Result = Result + WorksheetFunction.Median(WorksheetFunction.MOD(endTime, 1), ShiftEndTime, ShiftStartTime)
'Else
'    Result = Result + ShiftEndTime
'End If
'
'Result = Result - WorksheetFunction.Median(WorksheetFunction.NetworkDays(receivedTime, receivedTime) * WorksheetFunction.MOD(receivedTime, 1), ShiftEndTime, ShiftStartTime)
'---------------------------------------------------
        
End Function
