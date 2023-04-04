Attribute VB_Name = "Module6"
Option Explicit
Sub sauv_geo(chemin, nom_fich)

Dim fichierk_x As String
Dim msg, style, title, response As String
Dim i, dec As Integer
Dim ecriture As Boolean

fichierk_x = chemin & nom_fich & ".geo"
ecriture = True


'Effacement des fichiers précédents

If Dir$(fichierk_x) <> "" Then
'Demander confirmation de l'effacement du fichier
msg = "Fichier déjà existant, Voulez vous l'écraser ?" ' Définit le message.
style = vbYesNo + vbDefaultButton2  ' Définit les boutons.
title = "Ecraser ?" ' Définit le titre.
response = MsgBox(msg, style, title)
If response = "6" Then
Kill fichierk_x
Else
ecriture = False
End If

End If


If ecriture = True Then
'Remplissage du fichier

Open fichierk_x For Append As #1
'Caractéristiques du bateau
Print #1, Range("'Données Générales'!B3").Value
Print #1, Range("'Données Générales'!B4").Value
Print #1, Range("'Données Générales'!B5").Value
Print #1, Range("'Données Générales'!B8").Value
Print #1, Range("'Données Générales'!B10").Value
Print #1, Range("'Données Générales'!B11").Value
Print #1, Range("'Données Générales'!B12").Value
Print #1, Range("'Données Générales'!B13").Value

'Fonction F1
Print #1, Range("'Données Générales'!F6").Value
Print #1, Range("'Données Générales'!G6").Value
Print #1, Range("'Données Générales'!F9").Value
Print #1, Range("'Données Générales'!G9").Value
Print #1, Range("'Données Générales'!F12").Value
Print #1, Range("'Données Générales'!G12").Value
Print #1, Range("'Données Générales'!F15").Value
Print #1, Range("'Données Générales'!G15").Value

'Fonction H1
Print #1, Range("'Données Générales'!K6").Value
Print #1, Range("'Données Générales'!L6").Value
Print #1, Range("'Données Générales'!K9").Value
Print #1, Range("'Données Générales'!L9").Value
Print #1, Range("'Données Générales'!K12").Value
Print #1, Range("'Données Générales'!L12").Value
Print #1, Range("'Données Générales'!K15").Value
Print #1, Range("'Données Générales'!L15").Value

'Noeuds Ai, y et z
dec = 3

For i = 0 To 1
'Print #1, Range("'Données Générales'!C" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!D" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!E" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!F" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!G" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!H" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!I" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!J" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!K" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!L" & 2 * i + dec).Value
'Print #1, Range("'Données Générales'!M" & 2 * i + dec).Value
Print #1, Range("'Données Générales'!P" & i + dec).Value
Print #1, Range("'Données Générales'!Q" & i + dec).Value
Print #1, Range("'Données Générales'!R" & i + dec).Value
Print #1, Range("'Données Générales'!S" & i + dec).Value

Print #1, Range("'Données Générales'!P" & i + dec + 5).Value
Print #1, Range("'Données Générales'!Q" & i + dec + 5).Value
Print #1, Range("'Données Générales'!R" & i + dec + 5).Value
Print #1, Range("'Données Générales'!S" & i + dec + 5).Value
Next i

Print #1, Range("'Données Générales'!R11").Value
Print #1, Range("'Données Générales'!R13").Value
Print #1, Range("'Données Générales'!R14").Value
Print #1, Range("'Données Générales'!S13").Value
Print #1, Range("'Données Générales'!S14").Value

Close #1
End If


End Sub
Sub lire_geo(chemin, nom_fich)

Dim fichierk_x As String
Dim msg, style, title, response As String
Dim i, dec As Integer
Dim chaine As String
Dim x(120) As Single

Dim ecriture As Boolean



fichierk_x = chemin & nom_fich & ".geo"

ecriture = True
i = 0
Open fichierk_x For Input As #1


Do While Not EOF(1)
        Line Input #1, chaine
        x(i) = Val(chaine)
        i = i + 1
Loop


'Caractéristiques du bateau
Range("'Données Générales'!B3").Value = x(0)
Range("'Données Générales'!B4").Value = x(1)
Range("'Données Générales'!B5").Value = x(2)
Range("'Données Générales'!B8").Value = x(3)
Range("'Données Générales'!B10").Value = x(4)
Range("'Données Générales'!B11").Value = x(5)
Range("'Données Générales'!B12").Value = x(6)
Range("'Données Générales'!B13").Value = x(7)
 

'Fonction F1
Range("'Données Générales'!F6").Value = x(8)
Range("'Données Générales'!G6").Value = x(9)
Range("'Données Générales'!F9").Value = x(10)
Range("'Données Générales'!G9").Value = x(11)
Range("'Données Générales'!F12").Value = x(12)
Range("'Données Générales'!G12").Value = x(13)
Range("'Données Générales'!F15").Value = x(14)
Range("'Données Générales'!G15").Value = x(15)

'Fonction H1
Range("'Données Générales'!K6").Value = x(16)
Range("'Données Générales'!L6").Value = x(17)
Range("'Données Générales'!K9").Value = x(18)
Range("'Données Générales'!L9").Value = x(19)
Range("'Données Générales'!K12").Value = x(20)
Range("'Données Générales'!L12").Value = x(21)
Range("'Données Générales'!K15").Value = x(22)
Range("'Données Générales'!L15").Value = x(23)

'Noeuds Ai, y et z
dec = 3

For i = 0 To 1

'Range("'Données Générales'!C" & 2 * i + dec).Value = x(24 + 11 * i)
Range("'Données Générales'!P" & i + dec).Value = x(24 + 8 * i)
Range("'Données Générales'!Q" & i + dec).Value = x(25 + 8 * i)
Range("'Données Générales'!R" & i + dec).Value = x(26 + 8 * i)
Range("'Données Générales'!S" & i + dec).Value = x(27 + 8 * i)

Range("'Données Générales'!P" & i + dec + 5).Value = x(28 + 8 * i)
Range("'Données Générales'!Q" & i + dec + 5).Value = x(29 + 8 * i)
Range("'Données Générales'!R" & i + dec + 5).Value = x(30 + 8 * i)
Range("'Données Générales'!S" & i + dec + 5).Value = x(31 + 8 * i)



'Range("'Données Générales'!D" & 2 * i + dec).Value = x(25 + 11 * i)
'Range("'Données Générales'!E" & 2 * i + dec).Value = x(26 + 11 * i)
'Range("'Données Générales'!F" & 2 * i + dec).Value = x(27 + 11 * i)
'Range("'Données Générales'!G" & 2 * i + dec).Value = x(28 + 11 * i)
'Range("'Données Générales'!H" & 2 * i + dec).Value = x(29 + 11 * i)
'Range("'Données Générales'!I" & 2 * i + dec).Value = x(30 + 11 * i)
'Range("'Données Générales'!J" & 2 * i + dec).Value = x(31 + 11 * i)
'Range("'Données Générales'!K" & 2 * i + dec).Value = x(32 + 11 * i)
'Range("'Données Générales'!L" & 2 * i + dec).Value = x(33 + 11 * i)
'Range("'Données Générales'!M" & 2 * i + dec).Value = x(34 + 11 * i)

Next i

Range("'Données Générales'!R11").Value = x(40)
If x(41) = 0 Then x(41) = 0.01
If x(43) = 0 Then x(43) = 0.01
Range("'Données Générales'!R13").Value = x(41)
Range("'Données Générales'!R14").Value = x(42)
Range("'Données Générales'!S13").Value = x(43)
Range("'Données Générales'!S14").Value = x(44)

Close #1



End Sub
