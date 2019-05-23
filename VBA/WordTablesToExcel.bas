Attribute VB_Name = "Module3"


Sub ImportWordTableByLine()

Dim r&
Dim strFile$, strFolder$
Dim ws As Object
Dim wdDoc As Object
Dim wdFileName As Variant
Dim TableNo As Integer 'table number in Word
Dim iRow As Long 'row index in Word
Dim jRow As Integer 'row index in Excel
Dim iCol As Integer 'column index in Word
Dim jCol As Integer 'column index in Excel

Dim lastrow As Long

'Sélectionner un dossier
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select Location Directory"
    .ButtonName = "Open"
    If .Show = -1 Then
        strFolder = .SelectedItems.Item(1) & "\"
    Else
        MsgBox "Action Canceled"
    End If
End With

'Sélectionne le premier fichier
strFile = Dir(strFolder) '//First file
jRow = 1

While Not strFile = ""
    jRow = jRow + 1
    jCol = 1
    strFile = strFolder + strFile
    
    'Ouvre le fichier Word
    Set wdDoc = GetObject(strFile)
        
        With wdDoc
            'Si il n'y a pas de tableau, message d'erreur
            If wdDoc.Tables.Count = 0 Then
                MsgBox "This document contains no tables", _
                    vbExclamation, "Import Word Table"
            Else
                'Passe sur la feuille Données
                Set ws = Worksheets("Données")
                Sheets("Données").Select
                
                'Parcours les tableaux un a un et copie les valeurs dans des cellules sur une même ligne
                For TableNo = 1 To wdDoc.Tables.Count
                    With .Tables(TableNo)
                        For iRow = 1 To .Rows.Count
                            For iCol = 1 To .Columns.Count
                                On Error Resume Next
                                'Copie des valeurs
                                ActiveSheet.Cells(jRow, jCol) = WorksheetFunction.Clean(.Cell(iRow, iCol).Range.Text)
                                jCol = jCol + 1
                                On Error GoTo 0
                            Next iCol
                        Next iRow
                    End With
                Next TableNo
            End If
        End With
        
    Set wdDoc = Nothing
    
    strFile = Dir() 'Passe au fichier suivant

Wend
MsgBox "Complete"
End Sub



