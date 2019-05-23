Attribute VB_Name = "Module2"


Sub ImportWordInvoice()

Dim r&
Dim strFile$, strFolder$
Dim ws As Object
Dim wdDoc As Object
Dim wdFileName As Variant
Dim TableNo As Integer 'table number in Word
Dim iRow As Long 'row index in Word
Dim jRow As Long 'row index in Excel
Dim iCol As Integer 'column index in Excel
Dim lastrow As Long


With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select Location Directory"
    .ButtonName = "Open"
    If .Show = -1 Then
    strFolder = .SelectedItems.Item(1) & "\"
    Else
    MsgBox "Action Canceled"
    End If
End With

'strFolder = "C:\Users\Johnny BeGood\Desktop\Nouveau dossier"
strFile = Dir(strFolder) '//First file
    
    
    While Not strFile = ""
        strFile = strFolder + strFile

        Set wdDoc = GetObject(strFile)

'write filename to static cell
Sheets("Feuil1").Select
Cells(2, 9) = strFile



With wdDoc

    Sheets.Add After:=Sheets(Sheets.Count) 'creates a new worksheet
    Sheets(Sheets.Count).Name = wdDoc.Name ' renames the new worksheet
    If wdDoc.Tables.Count = 0 Then
        MsgBox "This document contains no tables", _
            vbExclamation, "Import Word Table"
    Else
        jRow = 0
       Set ws = Worksheets("Feuil2")
       Sheets(wdDoc.Name).Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
        For TableNo = 1 To wdDoc.Tables.Count
            With .Tables(TableNo)
'copy cell contents from Word table cells to Excel cells
                For iRow = 1 To .Rows.Count
                    jRow = jRow + 1
                    For iCol = 1 To .Columns.Count
                        On Error Resume Next
                        ActiveSheet.Cells(jRow, iCol) = WorksheetFunction.Clean(.Cell(iRow, iCol).Range.Text)
                        On Error GoTo 0
                    Next iCol
                Next iRow
            End With
            jRow = jRow + 1
            
        Next TableNo
    End If
End With


 
'Çopy and paste selection as values in last row of GL sheet
        Sheets("Feuil1").Range("A2:J2").Copy
        Sheets("Feuil2").Activate
lastrow = Range("A65536").End(xlUp).Row
Sheets("Feuil2").Activate
Cells(lastrow + 1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Set wdDoc = Nothing

        strFile = Dir() '// Fetch next file in a folder

    Wend
MsgBox "Complete"
End Sub

