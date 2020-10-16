Attribute VB_Name = "Module1"
Sub sheetExtract()
'module
Dim ws As Worksheet
Dim x As Integer
Dim WS_Count As Integer
Dim name_sheet As String
Dim wb As Workbook
Dim wb_name As String ' nom du classeur
Dim timestamp As String
Dim folder As String 'dossier où on enregistre
timestamp = Format(Now(), "dd-MM-yyyy hh-mm-ss") 'on récupère la date et l'heure pour éviter d'écraser dossier
'MsgBox ThisWorkbook.Path & "\extraction\" & timestamp

folder = ThisWorkbook.Path & "\" & timestamp & " extraction\"
MkDir folder  'on crée le dossier dans lequel on va mettre les sheets"
WS_Count = ActiveWorkbook.Worksheets.Count
wb_name = ActiveWorkbook.Name


x = 1


For x = 3 To WS_Count
    Set wb = Workbooks.Add
    name_sheet = Workbooks(wb_name).Worksheets(x).Name
    'MsgBox ActiveWorkbook.Worksheets(x).Name
    Workbooks(wb_name).Sheets(name_sheet).Copy Before:=wb.Sheets(1)
    'wb.SaveAs "C:\temp8\" & name_sheet & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=folder & name_sheet, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    wb.Sheets("Feuil1").Delete
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=True
    
Next x






End Sub


