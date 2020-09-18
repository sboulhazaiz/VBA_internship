Sub sheetExtract()

Dim ws As Worksheet
Dim x As Integer
Dim WS_Count As Integer
Dim name_sheet As String
Dim wb As Workbook
Dim wb_name As String ' nom du classeur
Dim timestamp As String

timestamp = Format(Now(), "yyyy-MM-dd hh:mm:ss")
MsgBox timestamp
wb_name = ActiveWorkbook.Name
Set wb = Workbooks.Add
x = 1

'MkDir
'MsgBox wb_name

WS_Count = ActiveWorkbook.Worksheets.Count
 
For x = 1 To 0
    name_sheet = Workbooks("ERGOSUP_SYNTHESE_SEPTEMBRE 2020_OFFICIEL.xlsm").Worksheets(x).Name
    'MsgBox ActiveWorkbook.Worksheets(x).Name
    Workbooks("ERGOSUP_SYNTHESE_SEPTEMBRE 2020_OFFICIEL.xlsm").Sheets(name_sheet).Copy Before:=wb.Sheets(1)
    'wb.SaveAs "C:\temp8\" & name_sheet & ".xlsx"
    wb.SaveAs Filename:=ThisWorkbook.Path & "\extraction\" & name_sheet, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
Next x

'MsgBox x 'nombre de sheets, reste plus qu'à savoir y accéder singulièrement



End Sub

