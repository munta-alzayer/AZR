Attribute VB_Name = "Module1"
Sub MoveColumns()

Dim iRow As Long
Dim iCol As Long 'Constant values
data_sheet1 = ActiveSheet.Name
'Create Input Box to ask the user which sheet needs to be reorganised
target_sheet = "Final Report" 'Specify the sheet to store the results
iRow = Sheets(data_sheet1).UsedRange.Rows.Count
'Determine how many rows are in use
'Create a new sheet to store the results
Worksheets.Add.Name = "Final Report"

For iCol = 1 To Sheets(data_sheet1).UsedRange.Columns.Count

'Sets the TargetCol to zero in order to prevent overwriting existing targetcolumns
TargetCol = 0

'Read the header of the original sheet to determine the column order
If Sheets(data_sheet1).Cells(3, iCol).Value = "Top     " Then TargetCol = 1
If Sheets(data_sheet1).Cells(3, iCol).Value = "Bottom  " Then TargetCol = 2
If Sheets(data_sheet1).Cells(3, iCol).Value = "Length  " Then TargetCol = 3
If Sheets(data_sheet1).Cells(3, iCol).Value = "TNom    " Then TargetCol = 4
If Sheets(data_sheet1).Cells(3, iCol).Value = "TMin    " Then TargetCol = 5
If Sheets(data_sheet1).Cells(3, iCol).Value = "DptMxLos" Then TargetCol = 6
If Sheets(data_sheet1).Cells(3, iCol).Value = "MaxLoss%" Then TargetCol = 7


'If a TargetColumn was determined (based upon the header information) then copy the column to the right spot
If TargetCol <> 0 Then
'Select the column and copy it
Sheets(data_sheet1).Range(Sheets(data_sheet1).Cells(3, iCol), Sheets(data_sheet1).Cells(iRow, iCol)).Copy Destination:=Sheets(target_sheet).Cells(1, TargetCol)
End If

Next iCol 'Move to the next column until all columns are read

Range("A1:G1").Interior.Color = RGB(79, 98, 40)
Range("A1:G1").HorizontalAlignment = xlCenter
Range("A1:G1").VerticalAlignment = xlCenter
Range("A1:G1").WrapText = True
Range("A1:G1").Font.Color = vbWhite
Range("A1:G1").Font.Size = 9
Columns("A:G").ColumnWidth = 9

For currCol = 1 To Sheets(target_sheet).UsedRange.Columns.Count
    DecPlcs = 1
    If ActiveSheet.Cells(1, currCol).Value = "Top     " Then
        DecPlcs = 0
        for_mat = "###0"
        ActiveSheet.Cells(1, currCol).Value = "Top Depth(ft)"
    End If

    If ActiveSheet.Cells(1, currCol).Value = "Bottom  " Then
        DecPlcs = 0
        for_mat = "###0"
        ActiveSheet.Cells(1, currCol).Value = "Bottom Depth(ft)"
    End If

    If ActiveSheet.Cells(1, currCol).Value = "Length  " Then
        DecPlcs = 0
        for_mat = "###0"
        ActiveSheet.Cells(1, currCol).Value = "Body Length(ft)"
    End If
    If ActiveSheet.Cells(1, currCol).Value = "TNom    " Then
        DecPlcs = 3
        for_mat = "0.000"
        ActiveSheet.Cells(1, currCol).Value = "NomThk(in)"
    End If

    If ActiveSheet.Cells(1, currCol).Value = "TMin    " Then
        DecPlcs = 3
        for_mat = "0.000"
        ActiveSheet.Cells(1, currCol).Value = "MinThk(in)"
    End If

    If ActiveSheet.Cells(1, currCol).Value = "DptMxLos" Then
        DecPlcs = 0
        for_mat = "###0"
        ActiveSheet.Cells(1, currCol).Value = "MaxWL Depth(ft)"
        
    End If

    If ActiveSheet.Cells(1, currCol).Value = "MaxLoss%" Then
        DecPlcs = 1
        for_mat = "#.0"
        ActiveSheet.Cells(1, currCol).Value = "MaxWL(%)"
    End If
    
    ActiveSheet.Cells(1, currCol).Borders.LineStyle = xlContinous
    For currRow = 2 To Sheets(target_sheet).UsedRange.Rows.Count
        ActiveSheet.Cells(currRow, currCol).Value = Application.Round(ActiveSheet.Cells(currRow, currCol).Value, DecPlcs)
        ActiveSheet.Cells(currRow, currCol).NumberFormat = for_mat
    Next currRow
Next currCol


End Sub
