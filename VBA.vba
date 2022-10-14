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

If ShiftStartTimeString = "" Or ShiftEndTimeString = "" Or receivedTimeString = "" Then
    slaCalculate = "-"
    Exit Function
End If

receivedTime = CDate(receivedTimeString)
ShiftStartTime = CDate(ShiftStartTimeString)
ShiftEndTime = CDate(ShiftEndTimeString)

'If received time is in non-working hrs make it to start of next working hrs or monday morning if received in weekend
If Weekday(receivedTime, vbSaturday) = 1 Or Weekday(receivedTime, vbSunday) = 1 Or TimeSerial(Hour:=Hour(receivedTime), Minute:=Minute(receivedTime), Second:=Second(receivedTime)) > TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime)) Then
    If Weekday(receivedTime, vbFriday) = 1 Then
        receivedTime = DateAdd("d", 3, receivedTime)
    ElseIf Weekday(receivedTime, vbSaturday) = 1 Then
        receivedTime = DateAdd("d", 2, receivedTime)
    Else
        receivedTime = DateAdd("d", 1, receivedTime)
    End If
    receivedTime = CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftStartTime), Minute:=Minute(ShiftStartTime), Second:=Second(ShiftStartTime)))
ElseIf Weekday(receivedTime, vbSaturday) = 1 Or Weekday(receivedTime, vbSunday) = 1 Or TimeSerial(Hour:=Hour(receivedTime), Minute:=Minute(receivedTime), Second:=Second(receivedTime)) < TimeSerial(Hour:=Hour(ShiftStartTime), Minute:=Minute(ShiftStartTime), Second:=Second(ShiftStartTime)) Then
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

'For calculating weekends
If supportDays = 7 Then
    WeekendHrs = 0
    tempHrCal2 = slaTimeLimit
ElseIf supportDays = 5 Then
    'Calculating hours to add with non-working hours
    tempHrCal = DateDiff("n", receivedTime, CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime))))
    tempHrCal = (slaTimeLimit / 60) - (tempHrCal / 60)
    If tempHrCal > 0 Then
        tempHrCal2 = WorksheetFunction.RoundUp(tempHrCal / (DateDiff("n", ShiftStartTime, ShiftEndTime) / 60), 0)
        tempHrCal2 = tempHrCal2 * (24 - (DateDiff("n", ShiftStartTime, ShiftEndTime) / 60))
        tempHrCal2 = Abs((tempHrCal2 * 60) + slaTimeLimit)
    Else
        tempHrCal2 = slaTimeLimit
    End If

    If Weekday(receivedTime, vbFriday) = 1 Then
        FridayDate = CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime)))
        If receivedTime > FridayDate Then
            FridayDate = DateAdd("d", 8 - Weekday(receivedTime, vbFriday), receivedTime)
        Else
            FridayDate = receivedTime
        End If
    Else
        FridayDate = DateAdd("d", 8 - Weekday(receivedTime, vbFriday), receivedTime)
    End If
    FridayDate = CDate(DateSerial(Year:=Year(FridayDate), Month:=Month(FridayDate), Day:=Day(FridayDate)) & " " & TimeSerial(Hour:=Hour(ShiftEndTime), Minute:=Minute(ShiftEndTime), Second:=Second(ShiftEndTime)))
   
   ' WeekendHrs = (TimeSerial(24, 0, 0)) * (DateDiff("h", ShiftStartTime, ShiftEndTime)) - WorksheetFunction.NetworkDays(receivedTime, FridayDate)
    WeekendHrs = (DateDiff("n", ShiftStartTime, ShiftEndTime) / 60) * (DateDiff("d", receivedTime, FridayDate) + 1)
    'Subtracting hours from shift start time to received time
    WeekendHrs = WeekendHrs - Abs(DateDiff("n", receivedTime, CDate(DateSerial(Year:=Year(receivedTime), Month:=Month(receivedTime), Day:=Day(receivedTime)) & " " & TimeSerial(Hour:=Hour(ShiftStartTime), Minute:=Minute(ShiftStartTime), Second:=Second(ShiftStartTime)))) / 60)
    'Checking number of weekends it will take
    WeekendHrs = WorksheetFunction.RoundUp(((slaTimeLimit / 60) - WeekendHrs) / ((DateDiff("n", ShiftStartTime, ShiftEndTime) / 60) * 5), 0)
    'if remaining hours is less than friday shift end or (received date is friday and sla is less than working hours) then no weekend to be calculated
    If WeekendHrs < 0 Or (Weekday(receivedTime, vbFriday) = 1 And slaTimeLimit <= DateDiff("n", receivedTime, FridayDate)) Then WeekendHrs = 0
    WeekendHrs = WeekendHrs * 48
End If

'Answer = receivedtime + (total hours including non working hrs) + weekends if any
slaCalculate = DateAdd("n", tempHrCal2, receivedTime) + TimeSerial(WeekendHrs, 0, 0)

        
End Function

