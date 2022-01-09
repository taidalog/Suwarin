Attribute VB_Name = "main"
Option Explicit

Public Sub MakeSeatingChart()
    
    Dim ST As Double: ST = Timer
    
    Dim firstBorderedCell As Range
    Set firstBorderedCell = GetFirstBorderedCell(ActiveSheet.UsedRange)
    
    Dim topLeftSeatRange As Range
    Set topLeftSeatRange = GetTopLeftSeatRange(firstBorderedCell)
    
    Dim seatingChartRange As Range
    Set seatingChartRange = GetSeatingChartRange(firstBorderedCell)
    
    Dim seats() As Range
    seats = GetSeats(topLeftSeatRange, seatingChartRange)
    
    Dim attendees As Variant
    attendees = GetAttendees(seatingChartRange)
    
    If UBound(attendees, 1) > UBound(seats, 1) * UBound(seats, 2) Then
        MsgBox "Too many people."
        Exit Sub
    End If
    
    Dim skipString As String
    skipString = "x"
    
    Dim maxAttendeesForEachLine() As Long
    maxAttendeesForEachLine = DecideSeatArrangement(seats, UBound(attendees, 1), UBound(attendees, 1), skipString)
    
    Call ClearSeatingChart(seats, skipString, True)
    Call PutAttendeesToSeats(attendees, seats, maxAttendeesForEachLine, skipString)
    
    Debug.Print Timer - ST
    
End Sub


Private Function GetFirstBorderedCell(search_range As Range) As Range
    
    Dim forFrom1 As Long, forTo1 As Long, forFrom2 As Long, forTo2 As Long
    forFrom1 = 1
    forTo1 = search_range.Columns.Count
    forFrom2 = 1
    forTo2 = search_range.Rows.Count
    
    Dim i As Long
    For i = forFrom1 To forTo1
        Dim j As Long
        For j = forFrom2 To forTo2
            With search_range.Cells(j, i)
                If .Borders(xlEdgeTop).LineStyle <> xlNone Then
                    If .Borders(xlEdgeLeft).LineStyle <> xlNone Then
                        Set GetFirstBorderedCell = search_range.Cells(j, i)
                        Exit Function
                    End If
                End If
            End With
        Next j
    Next i
    
End Function


Private Function GetTopLeftSeatRange(first_bordered_cell As Range) As Range
    
    Dim i As Long
    i = 0
    
    Do While first_bordered_cell.Offset(i, 0).Borders(xlEdgeBottom).LineStyle = xlNone
        i = i + 1
        If first_bordered_cell.Offset(i, 0).Row = ActiveSheet.Rows.Count Then Exit Do
    Loop
    
    Dim j As Long
    j = 0
    
    Do While first_bordered_cell.Offset(0, j).Borders(xlEdgeRight).LineStyle = xlNone
        j = j + 1
        If first_bordered_cell.Offset(0, j).Column = ActiveSheet.Columns.Count Then Exit Do
    Loop
    
    Set GetTopLeftSeatRange = first_bordered_cell.Resize(i + 1, j + 1)
    
End Function


Private Function GetSeatingChartRange(first_bordered_cell As Range) As Range
    
    Dim i As Long
    i = 0
    
    Do While first_bordered_cell.Offset(i + 1, 0).Borders(xlEdgeLeft).LineStyle <> xlNone
        i = i + 1
        If first_bordered_cell.Offset(i, 0).Row = ActiveSheet.Rows.Count Then Exit Do
    Loop
    
    Dim j As Long
    j = 0
    
    Do While first_bordered_cell.Offset(0, j + 1).Borders(xlEdgeTop).LineStyle <> xlNone
        j = j + 1
        If first_bordered_cell.Offset(0, j).Column = ActiveSheet.Columns.Count Then Exit Do
    Loop
    
    Set GetSeatingChartRange = first_bordered_cell.Resize(i + 1, j + 1)
    
End Function


Private Function GetSeats(top_left_seat_range As Range, seating_chart_range As Range) As Range()
    
    Dim seatHeight As Long
    seatHeight = top_left_seat_range.Rows.Count
    
    Dim seatWidth As Long
    seatWidth = top_left_seat_range.Columns.Count
    
    Dim chartRows As Long
    chartRows = seating_chart_range.Rows.Count / seatHeight
    
    Dim chartColumns As Long
    chartColumns = seating_chart_range.Columns.Count / seatWidth
    
    
    Dim results() As Range
    ReDim results(1 To chartRows, 1 To chartColumns)
    
    Dim y As Long
    For y = 1 To chartColumns
        Dim x As Long
        For x = 1 To chartRows
            Set results(x, y) = top_left_seat_range.Offset((x - 1) * 2, (y - 1) * 2)
        Next x
    Next y
    
    GetSeats = results
    
End Function


Private Function GetAttendees(seating_chart_range As Range) As Variant
    
    With seating_chart_range
        Dim topRightCell As Range
        Set topRightCell = Intersect(.Item(1).EntireRow, .Item(.Count).EntireColumn).Offset(0, 2)
    End With
    
    GetAttendees = Range(topRightCell, topRightCell.End(xlDown)).Value
    
End Function


Private Function DecideSeatArrangement(seats_range() As Range, number_of_people As Long, number_of_needed_seats As Long, skip_string As String) As Long()
    
    Dim maxAttendeesForEachLine() As Long
    maxAttendeesForEachLine = DevideNumberEqually(number_of_needed_seats, UBound(seats_range(), 2), UBound(seats_range(), 1))
    
    Dim seatsToSkipCount As Long
    seatsToSkipCount = CountSeatsToSkip(seats_range, maxAttendeesForEachLine, skip_string)
    
    If number_of_needed_seats - seatsToSkipCount >= number_of_people Then
        DecideSeatArrangement = maxAttendeesForEachLine
    Else
        DecideSeatArrangement = DecideSeatArrangement(seats_range, number_of_people, number_of_people + seatsToSkipCount, skip_string)
    End If
    
End Function


Private Function DevideNumberEqually(number As Long, devide_into As Long, limit As Long) As Long()
    
    Dim results() As Long
    ReDim results(1 To devide_into)
    
    Dim i As Long
    For i = 1 To devide_into
        results(i) = Int(number / devide_into)
    Next i
    
    Dim remainingNumber As Long
    remainingNumber = number Mod devide_into
    
    If remainingNumber > 0 Then
        
        Dim numberToShift As Long
        numberToShift = Int((devide_into - remainingNumber) / 2)
        
        Dim j As Long
        For j = 1 + numberToShift To remainingNumber + numberToShift
            results(j) = results(j) + 1
        Next j
        
    End If
    
    DevideNumberEqually = results()
    
End Function


Private Function CountSeatsToSkip(seats_range() As Range, max_attendees_for_each_line() As Long, skip_string As String) As Long
    
    Dim seatsToSkipCount As Long
    seatsToSkipCount = 0
    
    Dim j As Long
    For j = 1 To UBound(seats_range, 2)
        Dim i As Long
        For i = 1 To max_attendees_for_each_line(j)
            If seats_range(i, j).Cells(1, 1).Value = skip_string Then seatsToSkipCount = seatsToSkipCount + 1
        Next i
    Next j
    
    CountSeatsToSkip = seatsToSkipCount
    
End Function


Private Sub ClearSeatingChart(seats_range() As Range, skip_string As String, leave_x As Boolean)
    
    Dim ST As Variant
    For Each ST In seats_range
        If ST.Cells(1, 1).Value <> skip_string Or leave_x = False Then
            ST.Cells(1, 1).ClearContents
        End If
    Next ST
    
End Sub


Private Sub PutAttendeesToSeats(attendees_array As Variant, seats_range() As Range, max_attendees_for_each_line() As Long, skip_string As String)
    
    Dim n As Long
    n = 1
    
    Dim j As Long
    For j = 1 To UBound(seats_range, 2)
        Dim i As Long
        For i = 1 To max_attendees_for_each_line(j)
            With seats_range(i, j).Cells(1, 1)
                If .Value <> skip_string Then
                    .Value = attendees_array(n, 1)
                    n = n + 1
                End If
            End With
'            If n > UBound(attendees_array, 1) Then Exit For
        Next i
'        If n > UBound(attendees_array, 1) Then Exit For
    Next j
    
End Sub