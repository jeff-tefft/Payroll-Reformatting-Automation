Sub Payroll_Format()
Dim main_workbook As Workbook
Dim get_data_from_file As String
Dim source_workbook_name As Variant
Dim source_workbook As Workbook
Dim source_worksheet As Worksheet
Dim template_worksheet As Worksheet
Dim output_worksheet As Worksheet
Dim output_rows As Integer
Dim raw_rows As Integer
Dim row_difference As Integer
Dim function_sheet As Worksheet
Dim set_for_delete As Integer
Dim current_department As String
Dim id_number As Variant
Dim process As Variant
Dim check_date As Variant
Dim check_voucher As Variant
Dim net_paid As Variant
Dim last_column As Integer
Dim insert_column As Integer
Dim column_header As String
Dim enter_column As Integer
Dim employee_name As String
Dim copy_amount As Single
Dim output_row As Integer
Dim column_string As String
Dim data_range As String

Dim debug_row As Integer


'get main workbook
EnableEvents = True
Set main_workbook = ThisWorkbook

'check the data source being used
get_data_from_file = MsgBox("Would you like to import data from another worksheet? Otherwise, the data in RAW will be used.", vbQuestion + vbYesNo + vbDefaultButton2, "Data Source")
    If get_data_from_file = vbYes Then
        source_workbook_name = Application.GetOpenFilename()

        If source_workbook_name = False Then
            'popup and exit, as no file was selected
            MsgBox ("No source file was selected. The program will now stop.")
            Exit Sub
        End If
        
        'delete old RAW sheet
        On Error GoTo ImportSheet 'skips if can't delete
        Application.DisplayAlerts = False
        main_workbook.Worksheets("RAW").Delete
        Application.DisplayAlerts = True
        
        'import sheet and save new RAW sheet
ImportSheet:
        On Error GoTo 0
        Application.ScreenUpdating = False
        Set source_workbook = Application.Workbooks.Open(source_workbook_name)
        Set source_worksheet = source_workbook.Worksheets(1)
        source_worksheet.Name = "RAW"
        source_worksheet.Copy After:=main_workbook.Sheets(2)
        source_workbook.Close (False)
        Application.ScreenUpdating = True
    End If

Set source_worksheet = main_workbook.Sheets(3)

'sets template for adjustments in headers and then copy for data input after
Set template_worksheet = main_workbook.Worksheets("Template")
'check number of lines in Template for outputting
output_rows = template_worksheet.Range("N" & template_worksheet.Rows.Count).End(xlUp).Row

'pull data from the "RAW" sheet
'check number of lines in RAW
Set function_sheet = main_workbook.Worksheets("FUNCTION")

'no need to pad with extra lines because counts blanks already
function_sheet.Range("B5") = "=COUNTIF(RAW!E3:E999,""*"")"
raw_rows = function_sheet.Range("B5").Value

'Add or delete lines from Output sheet
row_difference = output_rows - raw_rows
set_for_delete = output_rows - 3
If set_for_delete < 2 Then
    set_for_delete = 2
End If
If row_difference < 0 Then
    'add rows
    row_difference = row_difference * -1
    For i = 1 To row_difference
    template_worksheet.Rows(set_for_delete).EntireRow.Insert
    output_rows = output_rows + 1
    Next i
ElseIf row_difference > 2 Then
    'delete rows
    For i = 1 To row_difference
    template_worksheet.Rows(set_for_delete).EntireRow.Delete
    output_rows = output_rows - 1
    Next i
End If

'Go through each row by column and add column to Template sheet

'add formula to FUNCTION sheet for easy column lookup
function_sheet.Range("B3") = "=IFERROR(MATCH(B2,Template!$A$1:$ZZ$1,0),-1)"
function_sheet.Range("B6") = "=COUNTIF(Template!$A$1:$ZZ$1,""*"")"
function_sheet.Range("B9") = "=SUBSTITUTE(ADDRESS(1,B8, 4), ""1"", """")"

'start loop for checking each column and transferring data
last_column = source_worksheet.Cells(2, Columns.Count).End(xlToLeft).Column
last_column = last_column + 1 'necessary to avid off-by-one caused by merged cells
column_header = "unassigned"
insert_column = 8 ' sets default insert column jsut in case

For source_column = 11 To last_column
'get column header
column_header = source_worksheet.Cells(2, source_column).Value
If column_header <> "" Then 'allows skipping the merged cells and only including new
    'lookup using FUNCTION sheet
    function_sheet.Range("B2") = column_header
    enter_column = function_sheet.Range("B3").Value
    'if not found, add column and create proper headers
    If enter_column = -1 Then
        function_sheet.Range("B2") = source_worksheet.Cells(2, source_column - 2).Value
        insert_column = function_sheet.Range("B3").Value
        'insert column
        template_worksheet.Columns(insert_column + 1).EntireColumn.Insert
        'copy formatting
        template_worksheet.Columns(insert_column).Copy
        template_worksheet.Columns(insert_column + 1).PasteSpecial Paste:=xlFormats
        Application.CutCopyMode = False
        Range("A1").Select
        template_worksheet.Cells(1, insert_column + 1) = column_header
        function_sheet.Range("B8") = (insert_column + 1)
        column_string = function_sheet.Range("B9").Value
        template_worksheet.Cells(output_rows, set_for_delete + 1) = "=SUM(" & column_string & "3:" & column_string & (output_rows - 1) & ")"
    End If
'end loop
End If
Next source_column

'delete "Output" sheet and replace with blank template - AFTER checking headers to update template
On Error GoTo CopyTemplate 'skips if can't delete
Application.DisplayAlerts = False
main_workbook.Worksheets("Output").Delete
Application.DisplayAlerts = True
'copy "template" sheet
CopyTemplate:
On Error GoTo 0
template_worksheet.Copy After:=main_workbook.Sheets(1)
Set output_worksheet = main_workbook.Sheets(2)
output_worksheet.Name = "Output"


'transfer data
'make sure to handle departments as set until changed
current_department = "Unassigned"
'loop through each row until the column ends
raw_rows = source_worksheet.Range("E" & source_worksheet.Rows.Count).End(xlUp).Row
If raw_rows < 4 Then
raw_rows = 4
End If
output_row = 2
For source_row = 4 To raw_rows
    'check if issued check exists
    check_voucher = source_worksheet.Range("I" & source_row).Value
        If check_voucher <> "" Then
        'update department if needed
        If source_worksheet.Range("B" & source_row).Value <> "" Then
            current_department = source_worksheet.Range("B" & source_row).Value
        End If
        
        'remaining static data to variables
        'check for blank issues
        If source_worksheet.Range("E" & source_row).Value = "" Then
            'reuse all previous values except the net
            net_paid = source_worksheet.Range("J" & source_row).Value
        Else
            employee_name = source_worksheet.Range("E" & source_row).Value
            id_number = source_worksheet.Range("F" & source_row).Value
            process = source_worksheet.Range("G" & source_row).Value
            check_date = source_worksheet.Range("H" & source_row).Value
            net_paid = source_worksheet.Range("J" & source_row).Value
        End If
        
        
        'static data to output sheet
        output_worksheet.Range("B" & output_row) = employee_name
        output_worksheet.Range("A" & output_row) = current_department
        output_worksheet.Range("C" & output_row) = id_number
        output_worksheet.Range("D" & output_row) = process
        output_worksheet.Range("E" & output_row) = check_date
        output_worksheet.Range("F" & output_row) = check_voucher
        output_worksheet.Range("G" & output_row) = net_paid

        'loop through columns and copy over data
        For source_column = 11 To last_column
            'only transfer data if Amount in Row 3
            If source_worksheet.Cells(3, source_column).Value = "Amount" Then
                copy_amount = source_worksheet.Cells(source_row, source_column).Value
                'update lookup in function sheet
                column_header = source_worksheet.Cells(2, source_column - 1).Value
                function_sheet.Range("B2") = column_header
                enter_column = function_sheet.Range("B3").Value
                'use lookup to paste value into output sheet
                output_worksheet.Cells(output_row, enter_column) = copy_amount
            End If
            'end column loop
        Next source_column
        output_row = output_row + 1
    End If
'end rowloop
Next source_row


'update pivot table
function_sheet.Range("B8") = "=B6"
column_string = function_sheet.Range("B9").Value
data_range = "Output!$A$1:$" & column_string & "$" & (output_row - 1)
For Each pivot_table In output_worksheet.PivotTables
         pivot_table.ChangePivotCache ActiveWorkbook.PivotCaches.Create _
            (SourceType:=xlDatabase, SourceData:=output_worksheet.Range(data_range))
Next pivot_table
ThisWorkbook.RefreshAll

'move view to convenient location
output_worksheet.Range("A" & output_rows).Select

End Sub
