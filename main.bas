Attribute VB_Name = "main"
Option Explicit

Public Enum enumSearchDirection
    SearchByColumn
    SearchByRow
End Enum

Public Enum enumSeatStart
    BottomLeft
    BottomRight
    TopLeft
    TopRight
End Enum

Public Enum enumSeatDirection
    ByColumn
    ByRow
End Enum

Public Enum enumSeatAlignment
    ToCenter
    ToFirst
    ToLast
End Enum


Public Sub MakeSeatingChart()
    ' main
    
    Dim ST As Double: ST = Timer
    
    Dim firstBorderedCell As Range
    Set firstBorderedCell = GetFirstBorderedCell(ActiveSheet.UsedRange, SearchByColumn)
    
    If firstBorderedCell Is Nothing Then
'        MsgBox "Format Error:" & vbCrLf & _
               "First bordered cell could not be found." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "最初の罫線付きセルが見つかりませんでした。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    Dim topLeftSeatRange As Range
    Set topLeftSeatRange = GetTopLeftSeatRange(firstBorderedCell)
    
    If topLeftSeatRange Is Nothing Then
'        MsgBox "Format Error:" & vbCrLf & _
               "Top left seat could not be found." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "左上の座席が見つかりませんでした。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    Dim seatingChartRange As Range
    Set seatingChartRange = GetSeatingChartRange(firstBorderedCell)
    
    If seatingChartRange Is Nothing Then
'        MsgBox "Format Error:" & vbCrLf & _
               "Seating chart range could not be found." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "座席表が見つかりませんでした。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    If seatingChartRange.Columns.Count Mod topLeftSeatRange.Columns.Count <> 0 Then
'        MsgBox "Format Error:" & vbCrLf & _
               "Some columns (vertical lines of seats) have wrong number of cells." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "座席表の縦の列のセル数が異なります。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    If seatingChartRange.Rows.Count Mod topLeftSeatRange.Rows.Count <> 0 Then
'        MsgBox "Format Error:" & vbCrLf & _
               "Some rows (horizontal lines of seats) have wrong number of cells." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "座席表の横の列のセル数が異なります。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    Dim seats() As Range
    seats = GetSeats(topLeftSeatRange, seatingChartRange)
    
    ' Judging whether the dynamic array variable is assigned (-1 means "NOT assigned.").
    If (Not seats) = -1 Then
'        MsgBox "Format Error:" & vbCrLf & _
               "Seats could not be found." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "座席が見つかりませんでした。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    Dim participants As Variant
    participants = GetParticipants(seatingChartRange)
    
    If IsEmpty(participants) Then
'        MsgBox "Format Error:" & vbCrLf & _
               "Participants could not be found." & vbCrLf & _
               "See help and make it sure that the seating chart has the correct format."
        MsgBox "フォーマット エラー:" & vbCrLf & _
               "参加者が見つかりませんでした。" & vbCrLf & _
               "ヘルプを参照して、座席表のフォーマットが正しいか確認してください。"
        Exit Sub
    End If
    
    ' Judging whether number of participants exceeds the number of seats or not.
    If UBound(participants, 1) > UBound(seats, 1) * UBound(seats, 2) Then
'        MsgBox "Capacity Error." & vbCrLf & _
               "Participants exceeded seats." & vbCrLf & _
               "Expand the seating chart or reduce the number of the participants."
        MsgBox "キャパシティ エラー:" & vbCrLf & _
               "参加者の数が座席数を超えました。" & vbCrLf & _
               "座席数を増やすか、参加者を減らしてください。"
        Exit Sub
    End If
    
    Dim stringToSkip As String
    stringToSkip = "x"
    
    Dim maxParticipantsForEachLine() As Long
    maxParticipantsForEachLine = DecideSeatArrangement(seats, UBound(participants, 1), UBound(participants, 1), stringToSkip, ByColumn, ToCenter)
    
    ' Judging whether the dynamic array variable is assigned (-1 means "NOT assigned.").
    If (Not maxParticipantsForEachLine) = -1 Then
        Exit Sub
    End If
    
    Call ClearSeatingChart(seats, stringToSkip, True)
    Call PutParticipantsToSeats(participants, seats, maxParticipantsForEachLine, stringToSkip)
    
    Debug.Print Timer - ST
    
End Sub


Private Function GetFirstBorderedCell(search_range As Range, search_direction As enumSearchDirection) As Range
    
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


Private Function GetParticipants(seating_chart_range As Range) As Variant
    
    With seating_chart_range
        Dim topRightCell As Range
        Set topRightCell = Intersect(.Item(1).EntireRow, .Item(.Count).EntireColumn)
    End With
    
    Dim topParticipantsRange As Range
    Set topParticipantsRange = topRightCell.Offset(0, 2)
    
    If topParticipantsRange.Value = "" Then
        GetParticipants = Empty
    Else
        GetParticipants = Range(topParticipantsRange, topParticipantsRange.End(xlDown)).Value
    End If
    
End Function


Private Function DecideSeatArrangement( _
    seats_range() As Range, _
    participants_count As Long, _
    needed_seats_count As Long, _
    string_to_skip As String, _
    seat_direction As enumSeatDirection, _
    seat_alignment As enumSeatAlignment _
    ) As Long()
    
    If needed_seats_count > UBound(seats_range, 1) * UBound(seats_range, 2) Then
'        MsgBox "Capacity Error:" & vbCrLf & _
               "Number of needed seats exceeded existing seats." & vbCrLf & _
               "Expand the seating chart or reduce the number of '" & string_to_skip & "'."
        MsgBox "キャパシティ エラー:" & vbCrLf & _
               "必要な座席表が実際の座席数を超えました。" & vbCrLf & _
               "座席数を増やすか、'" & string_to_skip & "'を減らしてください。"
        Exit Function
    End If
    
    Dim maxParticipantsForEachLine() As Long
    maxParticipantsForEachLine = DevideNumberEqually(needed_seats_count, UBound(seats_range(), 2), UBound(seats_range(), 1))
    
    Dim seatsToSkipCount As Long
    seatsToSkipCount = CountSeatsToSkip(seats_range, maxParticipantsForEachLine, string_to_skip)
    
    If needed_seats_count - seatsToSkipCount >= participants_count Then
        DecideSeatArrangement = maxParticipantsForEachLine
    Else
        DecideSeatArrangement = DecideSeatArrangement(seats_range, participants_count, participants_count + seatsToSkipCount, string_to_skip, ByColumn, ToCenter)
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
        
        If Int(number / devide_into) + 1 > limit Then
'            MsgBox "Capacity Error:" & vbCrLf & _
                   "Exceeded the limit for a line." & vbCrLf & _
                   "Expand the seating chart or reduce the number of the participants."
            MsgBox "キャパシティ エラー:" & vbCrLf & _
                   "一列あたりの人数の上限を超えました。" & vbCrLf & _
                   "座席数を増やすか、参加者を減らしてください。"
            Exit Function
        End If
        
        Dim numberToShift As Long
        numberToShift = Int((devide_into - remainingNumber) / 2)
        
        Dim j As Long
        For j = 1 + numberToShift To remainingNumber + numberToShift
            results(j) = results(j) + 1
        Next j
        
    End If
    
    DevideNumberEqually = results()
    
End Function


Private Function CountSeatsToSkip(seats_range() As Range, max_participants_for_each_line() As Long, string_to_skip As String) As Long
    
    Dim seatsToSkipCount As Long
    seatsToSkipCount = 0
    
    Dim j As Long
    For j = 1 To UBound(seats_range, 2)
        Dim i As Long
        For i = 1 To max_participants_for_each_line(j)
            If seats_range(i, j).Cells(1, 1).Value = string_to_skip Then seatsToSkipCount = seatsToSkipCount + 1
        Next i
    Next j
    
    CountSeatsToSkip = seatsToSkipCount
    
End Function


Private Sub ClearSeatingChart(seats_range() As Range, string_to_skip As String, leave_string_to_skip As Boolean)
    
    Dim ST As Variant
    For Each ST In seats_range
        If ST.Cells(1, 1).Value <> string_to_skip Or leave_string_to_skip = False Then
            ST.Cells(1, 1).ClearContents
        End If
    Next ST
    
End Sub


Public Sub CallClearSeatingChart()
    
    Dim firstBorderedCell As Range
    Set firstBorderedCell = GetFirstBorderedCell(ActiveSheet.UsedRange, SearchByColumn)
    
    Dim topLeftSeatRange As Range
    Set topLeftSeatRange = GetTopLeftSeatRange(firstBorderedCell)
    
    Dim seatingChartRange As Range
    Set seatingChartRange = GetSeatingChartRange(firstBorderedCell)
    
    Dim seats() As Range
    seats = GetSeats(topLeftSeatRange, seatingChartRange)
    
    Dim stringToSkip As String
    stringToSkip = "x"
    
    Dim leaveStringToSkip As Boolean
'    leaveStringToSkip = MsgBox("Do you want to leave '" & stringToSkip & "'?", vbYesNo) = vbYes
    leaveStringToSkip = MsgBox("'" & stringToSkip & "'を残しますか？", vbYesNo) = vbYes
    
    Call ClearSeatingChart(seats, stringToSkip, leaveStringToSkip)
    
End Sub


Private Sub PutParticipantsToSeats(participants_array As Variant, seats_range() As Range, max_participants_for_each_line() As Long, string_to_skip As String)
    
    Dim n As Long
    n = 1
    
    Dim j As Long
    For j = 1 To UBound(seats_range, 2)
        Dim i As Long
        For i = 1 To max_participants_for_each_line(j)
            With seats_range(i, j).Cells(1, 1)
                If .Value <> string_to_skip Then
                    .Value = participants_array(n, 1)
                    n = n + 1
                End If
            End With
'            If n > UBound(participants_array, 1) Then Exit For
        Next i
'        If n > UBound(participants_array, 1) Then Exit For
    Next j
    
End Sub


Public Sub CopyActivesheet()
    
    Dim copyNumber As Long
    copyNumber = Application.InputBox("How many copy do you want?", Default:=1, Type:=1)
    
    Dim AWS As Worksheet
    Set AWS = ActiveSheet
    
    Dim i As Long
    For i = 1 To copyNumber
        AWS.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Next i
    
    AWS.Select
    
End Sub


Public Sub AddToContextMenu()
    
    With Application.CommandBars
        
        Dim i As Long
        For i = 1 To .Count
            
            If .Item(i).Name = "Cell" Then
                
                With .Item(i).Controls.Add(Type:=msoControlPopup, Temporary:=True)
                    .BeginGroup = True
'                    .Caption = "&" & ThisWorkbook.Name
                    .Caption = ThisWorkbook.Name & "(&" & Mid(ThisWorkbook.Name, 1, 1) & ")"
                    
                    With .Controls.Add
'                        .Caption = "&Make Seating Chart"
                        .Caption = "座席表を作成する(&M)"
                        .OnAction = ThisWorkbook.Name & "!" & "MakeSeatingChart"
                    End With
                    
                    With .Controls.Add
'                        .Caption = "&Clear Seating Chart"
                        .Caption = "座席表を消去する(&C)"
                        .OnAction = ThisWorkbook.Name & "!" & "CallClearSeatingChart"
                    End With
                    
                    With .Controls.Add
'                        .Caption = "Co&py This Worksheet"
                        .Caption = "このシートを複製する(&P)"
                        .OnAction = ThisWorkbook.Name & "!" & "CopyActivesheet"
                    End With
                    
                End With
                
            End If
            
        Next i
        
    End With
    
End Sub
