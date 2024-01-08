
Private Sub UserForm_Initialize()
' initialize: insert combobox items, for book name, activity type and time comboboxes

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet2")

Set Rng = ws.Range("A2:A1000")
For i = 1 To Rng.Rows.Count
    If Not (IsEmpty(Rng.Cells(i, 1))) And Not (Rng.Cells(i, 1) = ".") Then
        new_plan_user_form.book_combo_box.AddItem (Rng.Cells(i, 1))
    End If
Next i

Set Rng = ws.Range("E2:E10")
For i = 1 To Rng.Rows.Count
    If Not (IsEmpty(Rng.Cells(i, 1))) Then
        new_plan_user_form.activity_type_combo_box.AddItem (Rng.Cells(i, 1))
    End If
Next i

Set Rng = ws.Range("F2:F10")
For i = 1 To Rng.Rows.Count
    If Not (IsEmpty(Rng.Cells(i, 1))) Then
        new_plan_user_form.time_combo_box.AddItem (Rng.Cells(i, 1))
    End If
Next i
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub book_combo_box_Change()
' change title1 comboboxes

Dim ws As Worksheet
Dim choice As String
Dim start_point As Integer
Dim final_point As Integer

Set ws = ThisWorkbook.Sheets("Sheet2")
Set Rng = ws.Range("A2:B1000")
choice = new_plan_user_form.book_combo_box.Value

' find start point and final point for book name rows
For i = 1 To Rng.Rows.Count
    If StrComp(choice, Rng.Cells(i, 1), vbTextCompare) = 0 Then
        start_point = i
        For j = start_point + 1 To Rng.Rows.Count
            If Not (IsEmpty(Rng.Cells(j, 1))) Then
                final_point = j - 1
                Exit For
            End If
            If j = Rng.Rows.Count - 1 Then
                final_point = Rng.Rows.Count
            End If
        Next j
    End If
Next i

' delete previous items from title1 comboboxes
new_plan_user_form.title1_1_combo_box.Clear
new_plan_user_form.title1_2_combo_box.Clear

' add items to title1 comboboxes
For i = start_point To final_point
    If Not (IsEmpty(Rng.Cells(i, 2))) Then
        new_plan_user_form.title1_1_combo_box.AddItem (Rng.Cells(i, 2))
        new_plan_user_form.title1_2_combo_box.AddItem (Rng.Cells(i, 2))
    End If
Next i
new_plan_user_form.title1_1_combo_box.AddItem (ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value)

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub title1_1_combo_box_Change()
' change title2 comboboxes

Dim ws As Worksheet
Dim choice As String
Dim start_point As Integer
Dim final_point As Integer

Set ws = ThisWorkbook.Sheets("Sheet2")
Set Rng = ws.Range("B2:C1000")
choice = new_plan_user_form.title1_1_combo_box.Value

' find start point and final point for title1_1 name rows
For i = 1 To Rng.Rows.Count
    If StrComp(choice, Rng.Cells(i, 1), vbTextCompare) = 0 Then
        start_point = i
        For j = start_point + 1 To Rng.Rows.Count
            If Not (IsEmpty(Rng.Cells(j, 1))) Then
                final_point = j - 1
                Exit For
            End If
            If j = Rng.Rows.Count - 1 Then
                final_point = Rng.Rows.Count
            End If
        Next j
    End If
Next i

' delete previous items from title2 comboboxes
new_plan_user_form.title2_1_combo_box.Clear
new_plan_user_form.title2_2_combo_box.Clear

' add items to title2 comboboxes
If Not (StrComp(choice, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value, vbTextCompare) = 0) Then
    For i = start_point To final_point
        If Not (IsEmpty(Rng.Cells(i, 2))) Then
            new_plan_user_form.title2_1_combo_box.AddItem (Rng.Cells(i, 2))
            new_plan_user_form.title2_2_combo_box.AddItem (Rng.Cells(i, 2))
        End If
    Next i
    new_plan_user_form.title2_1_combo_box.AddItem (ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value)
End If

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub title1_2_combo_box_Change()
' change title2_2 combobox

Dim ws As Worksheet
Dim choice As String
Dim start_point As Integer
Dim final_point As Integer

Set ws = ThisWorkbook.Sheets("Sheet2")
Set Rng = ws.Range("B2:C1000")
choice = new_plan_user_form.title1_2_combo_box.Value

' find start point and final point for title1_2 name rows
For i = 1 To Rng.Rows.Count
    If StrComp(choice, Rng.Cells(i, 1), vbTextCompare) = 0 Then
        start_point = i
        For j = start_point + 1 To Rng.Rows.Count
            If Not (IsEmpty(Rng.Cells(j, 1))) Then
                final_point = j - 1
                Exit For
            End If
            If j = Rng.Rows.Count - 1 Then
                final_point = Rng.Rows.Count
            End If
        Next j
    End If
Next i

' add items to title2_2 combobox
If Not (StrComp(choice, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value, vbTextCompare) = 0) Then
    For i = start_point To final_point
        If Not (IsEmpty(Rng.Cells(i, 2))) Then
            new_plan_user_form.title2_2_combo_box.AddItem (Rng.Cells(i, 2))
        End If
    Next i
End If

' delete previous items from title2_2 combobox if title1_2 is "All"
If StrComp(choice, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value) = 0 Then
    new_plan_user_form.title2_2_combo_box.Clear
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub title2_1_combo_box_Change()
' change title3 comboboxes

Dim ws As Worksheet
Dim choice As String
Dim start_point As Integer
Dim final_point As Integer

Set ws = ThisWorkbook.Sheets("Sheet2")
Set Rng = ws.Range("C2:D1000")
choice = new_plan_user_form.title2_1_combo_box.Value

' find start point and final point for title2_1 name rows
For i = 1 To Rng.Rows.Count
    If StrComp(choice, Rng.Cells(i, 1), vbTextCompare) = 0 Then
        start_point = i
        For j = start_point + 1 To Rng.Rows.Count
            If Not (IsEmpty(Rng.Cells(j, 1))) Then
                final_point = j - 1
                Exit For
            End If
            If j = Rng.Rows.Count - 1 Then
                final_point = Rng.Rows.Count
            End If
        Next j
    End If
Next i

' delete previous items from title3 comboboxes
new_plan_user_form.title3_1_combo_box.Clear
new_plan_user_form.title3_2_combo_box.Clear

' add items to title3 comboboxes
If Not (StrComp(choice, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value, vbTextCompare) = 0) Then
    For i = start_point To final_point
        If Not (IsEmpty(Rng.Cells(i, 2))) Then
            new_plan_user_form.title3_1_combo_box.AddItem (Rng.Cells(i, 2))
            new_plan_user_form.title3_2_combo_box.AddItem (Rng.Cells(i, 2))
        End If
    Next i
    new_plan_user_form.title3_1_combo_box.AddItem (ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value)
End If

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub title2_2_combo_box_Change()
' change title3_2 combobox

Dim ws As Worksheet
Dim choice As String
Dim start_point As Integer
Dim final_point As Integer

Set ws = ThisWorkbook.Sheets("Sheet2")
Set Rng = ws.Range("C2:D1000")
choice = new_plan_user_form.title2_2_combo_box.Value

' find start point and final point for title2_2 name rows
For i = 1 To Rng.Rows.Count
    If StrComp(choice, Rng.Cells(i, 1), vbTextCompare) = 0 Then
        start_point = i
        For j = start_point + 1 To Rng.Rows.Count
            If Not (IsEmpty(Rng.Cells(j, 1))) Then
                final_point = j - 1
                Exit For
            End If
            If j = Rng.Rows.Count - 1 Then
                final_point = Rng.Rows.Count
            End If
        Next j
    End If
Next i

' add items to title3_2 combobox
If Not (StrComp(choice, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value, vbTextCompare) = 0) Then
    For i = start_point To final_point
        If Not (IsEmpty(Rng.Cells(i, 2))) Then
            new_plan_user_form.title3_2_combo_box.AddItem (Rng.Cells(i, 2))
        End If
    Next i
End If

' delete previous items from title3_2 combobox if title2_2 is "All"
If StrComp(choice, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value) = 0 Then
    new_plan_user_form.title3_2_combo_box.Clear
End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub create_plan_buttom_Click()
' create plan and edit visual options

Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim book_name As String
Dim title1_1 As String
Dim title1_2 As String
Dim title2_1 As String
Dim title2_2 As String
Dim title3_1 As String
Dim title3_2 As String
Dim activity_type As String
Dim time As String
Dim book_name_start_point As Integer
Dim book_name_final_point As Integer
Dim title1_1_start_point As Integer
Dim title1_1_final_point As Integer
Dim title1_2_start_point As Integer
Dim title1_2_final_point As Integer
Dim title2_1_start_point As Integer
Dim title2_1_final_point As Integer
Dim title2_2_start_point As Integer
Dim title2_2_final_point As Integer
Dim title3_1_start_point As Integer
Dim title3_1_final_point As Integer
Dim title3_2_start_point As Integer
Dim title3_2_final_point As Integer
Dim temporary_title1 As String
Dim temporary_title2 As String
Dim width As Integer
Dim one_cell As Integer
Dim two_cell As Integer


Set ws1 = ThisWorkbook.Sheets("Sheet2")
Set Rng = ws1.Range("A2:D1000")
book_name = new_plan_user_form.book_combo_box.Value
title1_1 = new_plan_user_form.title1_1_combo_box.Value
title1_2 = new_plan_user_form.title1_2_combo_box.Value
title2_1 = new_plan_user_form.title2_1_combo_box.Value
title2_2 = new_plan_user_form.title2_2_combo_box.Value
title3_1 = new_plan_user_form.title3_1_combo_box.Value
title3_2 = new_plan_user_form.title3_2_combo_box.Value
activity_type = new_plan_user_form.activity_type_combo_box.Value
time = new_plan_user_form.time_combo_box.Value

' find start point and final point for book name rows
For i = 1 To Rng.Rows.Count
    If StrComp(book_name, Rng.Cells(i, 1), vbTextCompare) = 0 Then
        book_name_start_point = i
        For j = book_name_start_point + 1 To Rng.Rows.Count
            If Not (IsEmpty(Rng.Cells(j, 1))) Then
                book_name_final_point = j - 1
                Exit For
            End If
            If j = Rng.Rows.Count - 1 Then
                book_name_final_point = Rng.Rows.Count
            End If
        Next j
    End If
Next i

' find start point and final point for title1 rows
For i = book_name_start_point To book_name_final_point
    If StrComp(title1_1, Rng.Cells(i, 2), vbTextCompare) = 0 Then
        title1_1_start_point = i
        For j = title1_1_start_point + 1 To book_name_final_point + 1
            If Not (IsEmpty(Rng.Cells(j, 2))) Then
                title1_1_final_point = j - 1
                Exit For
            End If
            If j = book_name_final_point - 1 Then
                title1_1_final_point = book_name_final_point
            End If
        Next j
    End If
Next i
If StrComp(title1_1, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value) = 0 Or title1_1 = "" Then
    title1_1_start_point = book_name_start_point
    title1_1_final_point = book_name_final_point
End If

For i = book_name_start_point To book_name_final_point
    If StrComp(title1_2, Rng.Cells(i, 2), vbTextCompare) = 0 Then
        title1_2_start_point = i
        For j = title1_2_start_point + 1 To book_name_final_point + 1
            If Not (IsEmpty(Rng.Cells(j, 2))) Then
                title1_2_final_point = j - 1
                Exit For
            End If
            If j = book_name_final_point - 1 Then
                title1_2_final_point = book_name_final_point
            End If
        Next j
    End If
Next i

If title1_2 = "" Then
    title1_2_start_point = title1_1_start_point
    title1_2_final_point = title1_1_final_point
End If

' find start point and final point for title2 rows
For i = title1_1_start_point To title1_1_final_point
    If StrComp(title2_1, Rng.Cells(i, 3), vbTextCompare) = 0 Then
        title2_1_start_point = i
        For j = title2_1_start_point + 1 To title1_1_final_point + 1
            If Not (IsEmpty(Rng.Cells(j, 3))) Then
                title2_1_final_point = j - 1
                Exit For
            End If
            If j = title1_1_final_point - 1 Then
                title2_1_final_point = title1_1_final_point
            End If
        Next j
    End If
Next i
If StrComp(title2_1, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value) = 0 Or title2_1 = "" Then
    title2_1_start_point = title1_1_start_point
    title2_1_final_point = title1_1_final_point
End If

For i = title1_2_start_point To title1_2_final_point
    If StrComp(title2_2, Rng.Cells(i, 3), vbTextCompare) = 0 Then
        title2_2_start_point = i
        For j = title2_2_start_point + 1 To title1_2_final_point + 1
            If Not (IsEmpty(Rng.Cells(j, 3))) Then
                title2_2_final_point = j - 1
                Exit For
            End If
            If j = title1_2_final_point - 1 Then
                title2_2_final_point = title1_2_final_point
            End If
        Next j
    End If
Next i

If title2_2 = "" Then
    title2_2_start_point = title2_1_start_point
    title2_2_final_point = title2_1_final_point
End If

' find start point and final point for title3 rows
For i = title2_1_start_point To title2_1_final_point
    If StrComp(title3_1, Rng.Cells(i, 4), vbTextCompare) = 0 Then
        title3_1_start_point = i
        For j = title3_1_start_point + 1 To title2_1_final_point + 1
            If Not (IsEmpty(Rng.Cells(j, 4))) Then
                title3_1_final_point = j - 1
                Exit For
            End If
            If j = title2_1_final_point - 1 Then
                title3_1_final_point = title2_1_final_point
            End If
        Next j
    End If
Next i
If StrComp(title3_1, ThisWorkbook.Worksheets("Sheet2").Cells(1, 7).Value) = 0 Or title3_1 = "" Then
    title3_1_start_point = title2_1_start_point
    title3_1_final_point = title2_1_final_point
End If

For i = title2_2_start_point To title2_2_final_point
    If StrComp(title3_2, Rng.Cells(i, 4), vbTextCompare) = 0 Then
        title3_2_start_point = i
        For j = title3_2_start_point + 1 To title2_2_final_point + 1
            If Not (IsEmpty(Rng.Cells(j, 4))) Then
                title3_2_final_point = j - 1
                Exit For
            End If
            If j = title2_2_final_point - 1 Then
                title3_2_final_point = title2_2_final_point
            End If
        Next j
    End If
Next i

If title3_2 = "" Then
    title3_2_start_point = title3_1_start_point
    title3_2_final_point = title3_1_final_point
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' create plan in cells
Set ws2 = ThisWorkbook.Sheets("Sheet1")
plan_start_row = ActiveCell.Row
plan_start_column = ActiveCell.Column

ws2.Cells(plan_start_row, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 1).Value
ws2.Cells(plan_start_row + 1, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 2).Value
ws2.Cells(plan_start_row + 2, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 3).Value
ws2.Cells(plan_start_row + 3, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 4).Value
ws2.Cells(plan_start_row + 4, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 5).Value
ws2.Cells(plan_start_row + 5, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 8).Value
ws2.Cells(plan_start_row + 6, plan_start_column).Value = ThisWorkbook.Worksheets("Sheet2").Cells(1, 6).Value
ws2.Cells(plan_start_row, plan_start_column + 1).Value = book_name
ws2.Cells(plan_start_row + 4, plan_start_column + 1).Value = activity_type
ws2.Cells(plan_start_row + 6, plan_start_column + 1).Value = time


For i = title3_1_start_point To title3_2_final_point
    ws2.Cells(plan_start_row + 3, plan_start_column + i - title3_1_start_point + 1).Value = Rng.Cells(i, 4)
    For j = 0 To 100
        If Not (IsEmpty(Rng.Cells(i - j, 3))) Then
            temporary_title2 = Rng.Cells(i - j, 3)
            ws2.Cells(plan_start_row + 2, plan_start_column + i - title3_1_start_point + 1).Value = temporary_title2
            Exit For
        End If
    Next j
    
    For u = 0 To 100
        If Not (IsEmpty(Rng.Cells(i - u, 2))) Then
            temporary_title1 = Rng.Cells(i - u, 2)
            ws2.Cells(plan_start_row + 1, plan_start_column + i - title3_1_start_point + 1).Value = temporary_title1
            Exit For
        End If
    Next u
Next i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' edit visual options
If title3_2_final_point = title3_1_start_point Then
    one_cell = 2
Else
    one_cell = 0
End If
If title3_2_final_point - title3_1_start_point = 1 Then
    two_cell = 2
Else
    two_cell = 0
End If
width = title3_2_final_point - title3_1_start_point + 2 + one_cell + two_cell
plan_final_row = plan_start_row + 6
plan_final_column = plan_start_column + width
Set Rng = ws2.Range(Cells(plan_start_row, plan_start_column), Cells(plan_final_row, plan_final_column))

''' Merges
Application.DisplayAlerts = False
Range(Rng.Cells(1, 2), Rng.Cells(1, width)).Merge

For i = 2 To width
    For j = 0 To width - 1 - two_cell
        If Not (Rng.Cells(2, i).Value = Rng.Cells(2, i + j)) Then
            Range(Rng.Cells(2, i), Rng.Cells(2, i + j - 1 + one_cell)).Merge
            Exit For
        End If
    Next j
Next i
For i = 2 To width
    For j = 0 To width - 1 - two_cell
        If Not (Rng.Cells(3, i).Value = Rng.Cells(3, i + j)) Then
            Range(Rng.Cells(3, i), Rng.Cells(3, i + j - 1 + one_cell)).Merge
            Exit For
        End If
    Next j
Next i

If one_cell > 0 Then
    Range(Rng.Cells(4, 2), Rng.Cells(4, 2 + one_cell)).Merge
End If

Range(Rng.Cells(5, 2), Rng.Cells(5, width)).Merge
Range(Rng.Cells(6, 2), Rng.Cells(6, width)).Merge
Range(Rng.Cells(7, 2), Rng.Cells(7, width)).Merge

' two cell condition
If two_cell > 0 Then
    Rng.Cells(4, 4).Value = Rng.Cells(4, 3).Value
    Rng.Cells(4, 3).ClearContents
    Range(Rng.Cells(4, 2), Rng.Cells(4, 3)).Merge
    Range(Rng.Cells(4, 4), Rng.Cells(4, 5)).Merge
    If Rng.Cells(3, 3).Value = "" Then
        Range(Rng.Cells(3, 2), Rng.Cells(3, 5)).Merge
    Else
        Rng.Cells(3, 4).Value = Rng.Cells(3, 3).Value
        Rng.Cells(3, 3).ClearContents
        Range(Rng.Cells(3, 2), Rng.Cells(3, 3)).Merge
        Range(Rng.Cells(3, 4), Rng.Cells(3, 5)).Merge
    End If
    If Rng.Cells(2, 3).Value = "" Then
        Range(Rng.Cells(2, 2), Rng.Cells(2, 5)).Merge
    Else
        Rng.Cells(2, 4).Value = Rng.Cells(2, 3).Value
        Rng.Cells(2, 3).ClearContents
        Range(Rng.Cells(2, 2), Rng.Cells(2, 3)).Merge
        Range(Rng.Cells(2, 4), Rng.Cells(2, 5)).Merge
    End If
        
End If

Application.DisplayAlerts = True

''' Fonts and Borders and Colors
Rng.Font.Name = "B Mitra"
Rng.Font.Size = 12
Rng.Font.Bold = True
Range(Rng.Cells(4, 2), Rng.Cells(4, width)).Font.Bold = False
Rng.Cells(6, 2).Font.Bold = False
Rng.HorizontalAlignment = xlCenter
Rng.VerticalAlignment = xlCenter

Rng.EntireColumn.AutoFit
Rng.EntireRow.AutoFit

Range(Rng.Cells(1, 1), Rng.Cells(7, width)).Borders.Weight = xlMedium

Range(Rng.Cells(1, 2), Rng.Cells(4, width)).Interior.Color = RGB(233, 234, 236)
Range(Rng.Cells(1, 2), Rng.Cells(4, width)).Font.Color = RGB(51, 54, 82)
Range(Rng.Cells(1, 1), Rng.Cells(7, 1)).Interior.Color = RGB(51, 54, 82)
Range(Rng.Cells(1, 1), Rng.Cells(7, 1)).Font.Color = RGB(233, 234, 236)
Rng.Cells(5, 2).Interior.Color = RGB(250, 208, 44)
Rng.Cells(5, 2).Font.Color = RGB(51, 54, 82)
Rng.Cells(7, 2).Interior.Color = RGB(250, 208, 44)
Rng.Cells(7, 2).Font.Color = RGB(51, 54, 82)
Rng.Cells(6, 2).Interior.Color = RGB(144, 173, 198)
Rng.Cells(6, 2).Font.Color = RGB(0, 0, 0)

Unload Me
End Sub

