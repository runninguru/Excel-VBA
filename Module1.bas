Attribute VB_Name = "Module1"
Option Explicit

'identify column letter from column number. Should work for any column, but it wont return an error if the new column letter is invalid.
 Function column_letter(ByVal ColumnNumber As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    column_letter = s
End Function

 Sub to_value()

    Dim rng As Range

    For Each rng In ActiveSheet.UsedRange

        If rng.HasFormula Then

            rng.Formula = rng.Value

        End If

    Next rng

End Sub
 Sub delete_all_columns_except_arr_string(ByVal wb As Workbook, ByVal ws As Worksheet, ByRef arr() As String)
    Dim found As Boolean
    Dim i As Integer
    Dim initial_wb As Workbook
    Dim initial_ws As Worksheet
    Set initial_wb = ActiveWorkbook
    Set initial_ws = ActiveSheet
    wb.Activate
    ws.Select
    found = False
    
    last_column = wb.Sheets(ws.Name).Cells(1, Columns.Count).End(xlToLeft).column
    
    i = last_column
    Do While i >= 1
        found = is_in_array_string(Cells(1, i).Value, arr)
        If found = False Then
            Columns(i).EntireColumn.Delete
        End If
        i = i - 1
    Loop
    initial_wb.Activate
    initial_ws.Select
End Sub

 Sub delete_all_columns_except(ByVal wb As Workbook, ByVal ws As Worksheet, ByRef arr() As Variant)
    Dim found As Boolean
    Dim i As Integer
    Dim initial_wb As Workbook
    Dim initial_ws As Worksheet
    Set initial_wb = ActiveWorkbook
    Set initial_ws = ActiveSheet
    wb.Activate
    ws.Select
    found = False
    
    last_column = wb.Sheets(ws.Name).Cells(1, Columns.Count).End(xlToLeft).column
    
    i = last_column
    Do While i >= 1
        found = is_in_array(Cells(1, i).Value, arr)
        If found = False Then
            Columns(i).EntireColumn.Delete
        End If
        i = i - 1
    Loop
    initial_wb.Activate
    initial_ws.Select
End Sub

Sub copy_row(ByVal wb As Workbook _
            , ByVal from_ws As Worksheet _
            , ByVal to_ws As Worksheet _
            , ByVal copy_last_column As Integer _
            , ByVal from_row As Integer _
            , ByVal paste_row As Integer)
    'paste_row is byRef because we update it here after we paste data.
    'from_row is ByVal because it gets iterated in the loop when this sub is called.
    wb.Activate
    
    from_ws.Select
    Range(Cells(from_row, 1), Cells(from_row, copy_last_column)).Copy
    to_ws.Select
            
    Range(Cells(paste_row, 1), Cells(paste_row, copy_last_column)).PasteSpecial Paste:=xlPasteFormats
    Range(Cells(paste_row, 1), Cells(paste_row, copy_last_column)).PasteSpecial Paste:=xlPasteValues
            
    'paste_row = wb.Sheets(to_ws.Name).Range("A1").End(xlDown).Row + 1
    from_ws.Select
End Sub

'Checks each header in the array to see if they exist in the selected workbook/worksheet. If any of them do not exist, display an error message and end macro execution entirely.
Sub continue_if_headers_exist(ByVal wb As Workbook, ByVal ws As Worksheet, ArrValues() As Variant)
    Dim column As Integer
    Dim last_filled_header As Integer
    Dim initial_wb As Workbook
    Dim initial_ws As Worksheet
    Dim in_array As Boolean
    wb.Activate
    ws.Select
    header_found = False
    current_array_header = False
    in_array = False
    Set initial_wb = ActiveWorkbook
    Set initial_ws = ActiveSheet
    
    last_filled_header = Cells(1, Columns.Count).End(xlToLeft).column
    
    For col = last_filled_header To 1 Step -1
        in_array = is_in_array(Cells(1, col).Value, ArrValues)
        If in_array = False Then
            Exit For
        End If
    Next col
    If in_array = False Then
        MsgBox "Error: One or more required columns does not exist. Macro Stopped. No Action."
        initial_wb.Activate
        initial_ws.Select
        End
    End If
    initial_wb.Activate
    initial_ws.Select
End Sub

'this doesn't work if there isn't any data in rows 2+ of column A
Sub add_formula_column(ByVal ws As Worksheet _
                    , ByVal wb As Workbook _
                    , ByVal column_number As Integer _
                    , ByVal new_header_name As String _
                    , ByVal formula_string As String _
                    , ByVal is_formula As Boolean)
    Dim initial_wb As Workbook
    Dim initial_ws As Worksheet
    Dim last_row As Integer
    
    Set initial_wb = ActiveWorkbook
    Set initial_ws = ActiveSheet

    last_row = wb.Sheets(ws.Name).Range("A1").End(xlDown).Row
    
    'select the chosen wb/ws just in case these are different from the currently active ones.
    wb.Activate
    ws.Select
    Cells(1, column_number).Select
    Cells(1, column_number).Value = new_header_name
    Cells(2, column_number).Select
    If is_formula = True Then
        wb.Sheets(ws.Name).Cells(2, column_number).Formula = formula_string
    Else
        wb.Sheets(ws.Name).Cells(2, column_number).Value = formula_string
    End If
    If last_row > 2 Then
        wb.Sheets(ws.Name).Cells(2, column_number).AutoFill Destination:=Range(Cells(2, column_number), Cells(last_row, column_number))
        Range(Cells(2, column_number), Cells(last_row, column_number)).Select
        Calculate
    End If
    
    'reset the active wb/ws just in case we were modifying a different wb/ws than the previously active ones.
    initial_wb.Activate
    initial_ws.Select
End Sub

'Only continues code execution if all sheets in the array exist in this workbook. Otherwise, display a message and stop the macro entirely.
'Do all of this entirely within this function.
'only known-tested by passing in an array/variant type.
'To Use:
'Dim Arr() As Variant
'Dim result As Boolean
'Arr = Array("Sheet1", "Sheet2", "Sheet3")
'result = ContinueIfAllSheetsExist(Arr)
'
'This is generally meant to be used at the beginning of a macro, before anything really impactful has happened.
Public Function ContinueIfAllSheetsExist(current_book As Workbook, ArrValues() As Variant) As Boolean
    Dim ws As Worksheet
    Dim exists As String
    Dim sheetExists As Boolean
    current_book.Activate
    
    exists = ""
    For i = 1 To UBound(ArrValues, 1)
        sheetExists = False
        'MsgBox "[" & arrValues(i) & "]" & vbNewLine & vbNewLine
        'loop through all tabs in current workbook
        For Each ws In Worksheets
            'MsgBox "[" & ws.Name & "]"
            If ws.Name = ArrValues(i) Then
                sheetExists = True
            End If
        Next ws
        If sheetExists = False Then
            exists = exists & "The sheet '" & ArrValues(i) & "' does not exist in the " & current_book.Name & " workbook. " & vbNewLine & vbNewLine
        End If
    Next i
    If Len(exists) > 0 Then
        MsgBox exists & "Please open or rename the above tab(s) as necessary and reopen the Macro. Macro Stopped. No Action."
        End
    End If
    ContinueIfSheetsExist = True
End Function

'input a string, return the column # matching the first header match (first column in row 1, starting from column A) it finds.
Public Function get_header_number_from_name(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal header_name As String) As Integer
    Dim initial_wb As Workbook
    Dim initial_ws As Worksheet
    Dim last_column, return_column As Integer
    
    return_column = -1 'using this column number will throw an error if the proper column is not found!
    
    'set the initial wb/ws just in case they are different from the selected ones.
    Set initial_wb = ActiveWorkbook
    Set initial_ws = ActiveSheet
    
    wb.Activate
    ws.Select
    
    last_column = ws.Cells(1, Columns.Count).End(xlToLeft).column
    
    For i = 1 To last_column
        If Cells(1, i).Value = header_name Then
            If return_column = -1 Then
                return_column = i
            End If
        End If
    Next i
    
    'reselect the initial wb/ws just in case they are different from selected ones.
    initial_wb.Activate
    initial_ws.Select
    
    get_header_number_from_name = return_column
    
End Function

'get a list of header names. return in range.
'in parent sub, assign this function to an array created without any defined dimensions, such as:
    'Dim arr() As Variant
'If you define dimensions in the array, this code wont work because you would need to have the code manually re-define the dimensions when adding the array elements.
Public Function GetHeaders(ByVal wb As Workbook, ByVal ws As Worksheet) As Variant
    Dim i As Integer
    Dim result As Integer
    Dim initial_ws As Worksheet
    Dim initial_wb As Workbook
    Dim last_column As Integer
    Dim last_column_letter As String
    Dim hdr() As Variant
    
    Set initial_wb = ActiveWorkbook
    Set initial_ws = ActiveSheet
    
    wb.Activate
    ws.Select
    last_column = Range("A1").End(xlToRight).column
    last_column_letter = column_letter(last_column)
    'Dim hdr As Range
    i = 1
    Cells(1, 1).Select
    result = ws.Range("A1").CurrentRegion.Columns.Count
    hdr = ws.Range("A1:" & last_column_letter & "1")
    
    initial_wb.Activate
    initial_ws.Select
    
    GetHeaders = hdr
End Function

Function is_in_array_string(valToBeFound As Variant, arr() As String) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError: 'array is empty
        For Each element In arr
            If element = valToBeFound Then
                is_in_array_string = True
                Exit Function
            End If
        Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    is_in_array_string = False
End Function

Function is_in_array(valToBeFound As Variant, arr() As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError: 'array is empty
        For Each element In arr
            If element = valToBeFound Then
                is_in_array = True
                Exit Function
            End If
        Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    is_in_array = False
End Function
Function file_exists(ByVal str_file_name As String) As Boolean
    Dim str_file_exists As String
    str_file_exists = Dir(str_file_name)
    
    If str_file_exists = "" Then
        file_exists = False
    Else
        file_exists = True
    End If
 
End Function
