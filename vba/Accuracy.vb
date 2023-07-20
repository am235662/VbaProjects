Public Sub Accuracy_data()
 
Dim copyfile As Workbook
'created variable as array with ()
Dim openfile As Variant

'to browse and select tracker


openfile = Application.GetOpenFilename(Title:="Select latest tracker", MultiSelect:=True)

 
 If VarType(openfile) = vbBoolean Then
 Exit Sub
 End If
  
 
Set copyfile = Workbooks.Open(openfile(1))
'Opened workbook
copyfile.Sheets(1).Select

ActiveSheet.Cells(1, 1).AutoFilter field:=53, Criteria1:="COMPLETED"
ActiveSheet.Cells(1, 1).AutoFilter field:=17, Criteria1:="<>0"

ActiveSheet.Range("$c$1").Select
Range(Selection, Selection.End(xlDown)).Copy

Workbooks(1).Activate
Sheets("Accuracy").Range("A1").PasteSpecial

Workbooks(2).Activate
Sheets(1).Range("$N$1:$s$1").Select
Range(Selection, Selection.End(xlDown)).Copy

Workbooks(1).Activate
Sheets("Accuracy").Range("b1").PasteSpecial

Workbooks(2).Activate
Sheets(1).Range("AL1").Select
Range(Selection, Selection.End(xlDown)).Copy

Workbooks(1).Activate
Sheets("Accuracy").Range("h1").PasteSpecial

Workbooks(2).Activate
Sheets(1).Range("AY1").Select
Range(Selection, Selection.End(xlDown)).Copy

Workbooks(1).Activate
Sheets("Accuracy").Range("I1").PasteSpecial

Application.DisplayAlerts = False 'disable alerts
Application.CutCopyMode = False 'Clear clipboard
copyfile.Close savechanges:=False 'Close the Master CLDC file in original form after removing filters or any other changes


End Sub


Sub accuracy_percentage()

Sheets("Productivity").Activate
ActiveSheet.Cells(1, 1).AutoFilter field:=12, Criteria1:="<>Directly Sent"

Sheets("Combined").Activate
Range("g3").Select
Range(Selection, Selection.End(xlDown)).ClearContents

Sheets("Combined").Activate
'to add new sheet
Sheets.Add
ActiveSheet.Name = "new"

Sheets("Productivity").Activate
Range("K1:M1").Select
Range(Selection, Selection.End(xlDown)).Copy

Sheets("new").Activate
Range("A1").PasteSpecial

Sheets("Combined").Activate
Range("$g$3").Formula = "=COUNTIF(new!a:a,A3)"

Range("g3").Select
Range(Selection, Selection.End(xlDown)).Copy
Range("g3").PasteSpecial xlPasteValues

Application.DisplayAlerts = False
Sheets("new").Delete
Application.DisplayAlerts = True 'to delete the new temporary sheet without any prompt

Sheets("Productivity").Select
ActiveSheet.AutoFilterMode = False
 
 Sheets("Combined").Select
 


End Sub

