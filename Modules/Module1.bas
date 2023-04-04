Attribute VB_Name = "Module1"


Option Explicit



Sub Echelle_graphique()
Attribute Echelle_graphique.VB_Description = "Macro enregistrée le 06/11/2003 par Vincent"
Attribute Echelle_graphique.VB_ProcData.VB_Invoke_Func = "r\n14"

Dim x, ymin, ymax As Single
Dim i As Integer


' Macro3 Macro
' Macro enregistrée le 06/11/2003 par Vincent
'

'Fonction F1

    ActiveSheet.ChartObjects("Graphique 8").Activate
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = 0
        .MaximumScale = Range("B3").Value
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = Range("B3").Value
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    
'Fonction H1
   
    ActiveSheet.ChartObjects("Graphique 9").Activate
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = 0
        .MaximumScale = Range("B3").Value
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = -Range("B3").Value / 2
        .MaximumScale = Range("B3").Value / 2
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    
    
    
'Fonctions G
For i = 10 To 21

    If Range("B4").Value > Range("B13").Value + Range("B10").Value Then
    
      x = Range("B4").Value * 1.25
      ymin = Range("B13").Value - Range("B4").Value * 1.25
      ymax = Range("B13").Value
    
    Else
    
    x = Range("B13").Value + Range("B10").Value
    ymin = -Range("B10")
    ymax = Range("B13")
    
    End If
        
  
    
    
    ActiveSheet.ChartObjects("Graphique " & i).Activate
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = 0
        .MaximumScale = x
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = ymin
        .MaximumScale = ymax
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    
Next i

End Sub

Sub exporter()
Attribute exporter.VB_ProcData.VB_Invoke_Func = "e\n14"

Dim x As Single
Dim i, j, k As Integer
Dim fichierk_x, fichierk_y, fichierk_z, fichierk As String

Dim col, chemin_txt, nom_fich_txt As String

chemin_txt = "C:\Documents and Settings\Vincent\Mes documents\maquette bateau\Texte\"
nom_fich_txt = "Spline"
' Creation des feuilles de splines

Sheets("Vide").Select

For i = 1 To 11

    Sheets.Add
    ActiveSheet.Name = "G" & 12 - i

Next i

'Copie des données


'Copie de la valeur de x
For j = 3 To 13
    Convert_colonne j, col
    
    Sheets("P(H1)").Select
    x = Range(col & "14").Value
    
        
    Sheets("G" & j - 2).Select
    
            For i = 1 To 11
            Range("A" & i).Value = x
            Next i
    
    Sheets("Parametrique").Select
    'y
    Convert_colonne 8 + j, col
    Range(col & "28:" & col & "38").Select
    Selection.Copy
    
    Sheets("G" & j - 2).Select
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Sheets("Parametrique").Select
    'z
    Range(col & "41:" & col & "51").Select
    Selection.Copy
    
    Sheets("G" & j - 2).Select
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Next j


' Spline d'étrave

Sheets("Vide").Select
Sheets.Add
    ActiveSheet.Name = "G12"
    
    Sheets("P(H1)").Select
    Range("M14").Select
    Selection.Copy
    
    Sheets("G12").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
  '  Range("A1").Value = Range("A1").Value * 1.00001
    
    Range("B1").Value = 0
    
    Sheets("P(F1)").Select
    Range("M14").Select
    Selection.Copy
    
    Sheets("G12").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("B2").Value = 0

    Sheets("Données Générales").Select
    Range("B12").Select
    Selection.Copy
    
    Sheets("G12").Select
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Sheets("Données Générales").Select
    Range("B13").Select
    Selection.Copy
    
    Sheets("G12").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False




'Effacement des fichiers précédents


For k = 1 To 12
Sheets("G" & k).Select

If k > 9 Then
fichierk_x = chemin_txt & nom_fich_txt & k & "_x.txt"
fichierk_y = chemin_txt & nom_fich_txt & k & "_y.txt"
fichierk_z = chemin_txt & nom_fich_txt & k & "_z.txt"

Else
fichierk_x = chemin_txt & nom_fich_txt & "0" & k & "_x.txt"
fichierk_y = chemin_txt & nom_fich_txt & "0" & k & "_y.txt"
fichierk_z = chemin_txt & nom_fich_txt & "0" & k & "_z.txt"
End If

If Dir$(fichierk_x) <> "" And Dir$(fichierk_y) <> "" And Dir$(fichierk_z) <> "" Then
Kill fichierk_x
Kill fichierk_y
Kill fichierk_z


End If

Next k

If Dir$(chemin_txt & "Donnees.txt") <> "" Then
Kill chemin_txt & "Donnees.txt"

End If

' Constitution des <Fichiers Textes> :




For k = 1 To 11
Sheets("G" & k).Select

If k > 9 Then
fichierk_x = chemin_txt & nom_fich_txt & k & "_x.txt"
fichierk_y = chemin_txt & nom_fich_txt & k & "_y.txt"
fichierk_z = chemin_txt & nom_fich_txt & k & "_z.txt"

Else
fichierk_x = chemin_txt & nom_fich_txt & "0" & k & "_x.txt"
fichierk_y = chemin_txt & nom_fich_txt & "0" & k & "_y.txt"
fichierk_z = chemin_txt & nom_fich_txt & "0" & k & "_z.txt"
End If

Open fichierk_x For Append As #1
Open fichierk_y For Append As #2
Open fichierk_z For Append As #3


For j = 1 To 11
Print #1, Sheets("G" & k).Range("A" & j).Value
Print #2, Sheets("G" & k).Range("B" & j).Value
Print #3, Sheets("G" & k).Range("C" & j).Value

Next j
Close #1
Close #2
Close #3

Next k

' fichier texte de spline  12

Sheets("G12").Select
fichierk_x = chemin_txt & "Spline12_x.txt"
fichierk_y = chemin_txt & "Spline12_y.txt"
fichierk_z = chemin_txt & "Spline12_z.txt"
Open fichierk_x For Append As #1
Open fichierk_y For Append As #2
Open fichierk_z For Append As #3

For j = 1 To 2
Print #1, Sheets("G12").Range("A" & j).Value
Print #2, Sheets("G12").Range("B" & j).Value
Print #3, Sheets("G12").Range("C" & j).Value


Next j
Close #1
Close #2
Close #3

' Effacement des feuilles annexes

Sheets(Array("G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G10", "G11", "G12")).Select
    Sheets("G1").Activate
    ActiveWindow.SelectedSheets.Delete
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Données Générales").Select

'Fichier texte Données  importantes

fichierk = chemin_txt & "Donnees.txt"
Open fichierk For Append As #1


Print #1, Sheets("Données Générales").Range("B13").Value
Print #1, Sheets("Données Générales").Range("B10").Value
Print #1, Sheets("Données Générales").Range("B4").Value
Print #1, Sheets("Données Générales").Range("B9").Value

Close #1


End Sub





Sub Convert_colonne(numero_colonne, col)


If numero_colonne = 1 Then
col = "A"
Else
If numero_colonne = 2 Then
col = "B"
Else
If numero_colonne = 3 Then
col = "C"
Else
If numero_colonne = 4 Then
col = "D"
Else
If numero_colonne = 5 Then
col = "E"
Else
If numero_colonne = 6 Then
col = "F"
Else
If numero_colonne = 7 Then
col = "G"
Else
If numero_colonne = 8 Then
col = "H"
Else
If numero_colonne = 9 Then
col = "I"
Else
If numero_colonne = 10 Then
col = "J"
Else
If numero_colonne = 11 Then
col = "K"
Else
If numero_colonne = 12 Then
col = "L"
Else
If numero_colonne = 13 Then
col = "M"
Else
If numero_colonne = 14 Then
col = "N"
Else
If numero_colonne = 15 Then
col = "O"
Else
If numero_colonne = 16 Then
col = "P"
Else
If numero_colonne = 17 Then
col = "Q"
Else
If numero_colonne = 18 Then
col = "R"
Else
If numero_colonne = 19 Then
col = "S"
Else
If numero_colonne = 20 Then
col = "T"
Else
If numero_colonne = 21 Then
col = "U"
Else
If numero_colonne = 22 Then
col = "V"
Else
If numero_colonne = 23 Then
col = "W"
Else
If numero_colonne = 24 Then
col = "X"
Else
If numero_colonne = 25 Then
col = "Y"
Else
If numero_colonne = 26 Then
col = "Z"
Else
If numero_colonne = 27 Then
col = "AA"
Else
If numero_colonne = 28 Then
col = "AB"
Else
If numero_colonne = 29 Then
col = "AC"
Else




End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub
