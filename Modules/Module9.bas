Attribute VB_Name = "Module9"
Option Explicit

Sub courbes_aires(section, num_angle, aire, val_angle)
Dim col As String


If num_angle = 0 Then
Range("'Aires'!A" & section + 6).Value = points4(section, 0, 0)
End If

Convert_colonne num_angle + 2, col
Range("'Aires'!" & col & 4).Value = val_angle * 180 / Pi
Range("'Aires'!" & col & section + 6).Value = aire

End Sub
Sub export_plans_forme()

' entrées : Couples,nu,nv, discrétisation fixe (10 courbes) => facilité de création de graph
' zmin, zmax, ymin, ymax
'courb(type,N°courbe,n°point,coordonnée)
'points3(s, nv + i, 1)
'sorties : courbes, => ss prog pour l'affichage de la feuille

Dim courb(1, 10, 400, 2), y(10), z(10) As Single
Dim i, j, s, discret As Integer
Dim ymax, zmin, zmax As Single
Dim col As String

nu = Val(F_chargmt.TextBox3.Value)
nv = Val(F_chargmt.TextBox4.Value)

discret = 8

ymax = Val(Range("'Données générales'!B4").Value)

zmin = -Val(Range("'Données générales'!B10").Value)
zmax = Val(Range("'Données générales'!B13").Value)

For i = 0 To discret
y(i) = i * ymax / discret
z(i) = i * (zmax - zmin) / discret + zmin
Next i




'Courbes issues des plans ZX

For i = 0 To discret
For s = 0 To nu

For j = nv To 2 * nv

If points3(s, j, 1) <= y(i) And points3(s, j + 1, 1) >= y(i) Then
courb(0, i, s, 0) = points3(s, 0, 0)
courb(0, i, s, 1) = y(i)

If points3(s, j, 1) = y(i) Then
courb(0, i, s, 2) = points3(s, j, 2)
ElseIf points3(s, j + 1, 1) = y(i) Then
courb(0, i, s, 2) = points3(s, j + 1, 2)
Else
courb(0, i, s, 2) = (points3(s, j, 2) + points3(s, j + 1, 2)) / 2
End If
End If

If points3(s, j, 2) <= z(i) And points3(s, j + 1, 2) >= z(i) Then
courb(1, i, s, 0) = points3(s, 0, 0)
courb(1, i, s, 2) = z(i)

If points3(s, j, 2) = z(i) Then
courb(1, i, s, 1) = points3(s, j, 1)
ElseIf points3(s, j + 1, 2) = z(i) Then
courb(1, i, s, 1) = points3(s, j + 1, 1)
Else
courb(1, i, s, 1) = (points3(s, j, 1) + points3(s, j + 1, 1)) / 2
End If
End If

Next j

Next s
Next i


'effacement de la feuille "formes"
                Sheets("formes").Select
'                Rows("4:150").Select
'                Selection.Delete Shift:=xlUp
                Rows("3:300").Select
                Selection.ClearContents
                Range("A1").Select
'effacement de la feuille "formes_2"
                Sheets("formes_2").Select
'                Rows("4:150").Select
'                Selection.Delete Shift:=xlUp
                Rows("3:300").Select
                Selection.ClearContents
                Range("A1").Select
Sheets("Plans_XZ").Select



'export
For i = 0 To discret
For s = 0 To nu

If i = 0 Then
Range("'formes'!A" & s + 3).Value = s
Range("'formes_2'!A" & s + 3).Value = s
End If

Convert_colonne 2 * i + 2, col
Range("'formes'!" & col & s + 3).Value = courb(0, i, s, 0)
Range("'formes_2'!" & col & s + 3).Value = courb(1, i, s, 0)

Convert_colonne 2 * i + 3, col
Range("'formes'!" & col & s + 3).Value = courb(0, i, s, 2)
Range("'formes_2'!" & col & s + 3).Value = courb(1, i, s, 1)
Next s
Next i

For s = 0 To nu

Range("'formes'!R" & s + 3).Value = points3(s, nv, 0)
Range("'formes'!S" & s + 3).Value = points3(s, nv, 2)
Next s

End Sub
Sub export_formes()
Dim i, j, s As Integer


End Sub
