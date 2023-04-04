Attribute VB_Name = "Module6"
Option Explicit
Sub sauv_geo(chemin, nom_fich)

Dim fichierk_x As String
Dim msg, style, title, response As String
Dim i, dec As Integer
Dim ecriture As Boolean

fichierk_x = chemin & nom_fich & ".geo"
ecriture = True


'Effacement des fichiers pr�c�dents

If Dir$(fichierk_x) <> "" Then
'Demander confirmation de l'effacement du fichier
msg = "Fichier d�j� existant, Voulez vous l'�craser ?" ' D�finit le message.
style = vbYesNo + vbDefaultButton2  ' D�finit les boutons.
title = "Ecraser ?" ' D�finit le titre.
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
'Caract�ristiques du bateau
Print #1, Range("'Donn�es G�n�rales'!B3").Value
Print #1, Range("'Donn�es G�n�rales'!B4").Value
Print #1, Range("'Donn�es G�n�rales'!B5").Value
Print #1, Range("'Donn�es G�n�rales'!B8").Value
Print #1, Range("'Donn�es G�n�rales'!B10").Value
Print #1, Range("'Donn�es G�n�rales'!B11").Value
Print #1, Range("'Donn�es G�n�rales'!B12").Value
Print #1, Range("'Donn�es G�n�rales'!B13").Value

'Fonction F1
Print #1, Range("'Donn�es G�n�rales'!F6").Value
Print #1, Range("'Donn�es G�n�rales'!G6").Value
Print #1, Range("'Donn�es G�n�rales'!F9").Value
Print #1, Range("'Donn�es G�n�rales'!G9").Value
Print #1, Range("'Donn�es G�n�rales'!F12").Value
Print #1, Range("'Donn�es G�n�rales'!G12").Value
Print #1, Range("'Donn�es G�n�rales'!F15").Value
Print #1, Range("'Donn�es G�n�rales'!G15").Value

'Fonction H1
Print #1, Range("'Donn�es G�n�rales'!K6").Value
Print #1, Range("'Donn�es G�n�rales'!L6").Value
Print #1, Range("'Donn�es G�n�rales'!K9").Value
Print #1, Range("'Donn�es G�n�rales'!L9").Value
Print #1, Range("'Donn�es G�n�rales'!K12").Value
Print #1, Range("'Donn�es G�n�rales'!L12").Value
Print #1, Range("'Donn�es G�n�rales'!K15").Value
Print #1, Range("'Donn�es G�n�rales'!L15").Value

'Noeuds Ai, y et z
dec = 3

For i = 0 To 1
'Print #1, Range("'Donn�es G�n�rales'!C" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!D" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!E" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!F" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!G" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!H" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!I" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!J" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!K" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!L" & 2 * i + dec).Value
'Print #1, Range("'Donn�es G�n�rales'!M" & 2 * i + dec).Value
Print #1, Range("'Donn�es G�n�rales'!P" & i + dec).Value
Print #1, Range("'Donn�es G�n�rales'!Q" & i + dec).Value
Print #1, Range("'Donn�es G�n�rales'!R" & i + dec).Value
Print #1, Range("'Donn�es G�n�rales'!S" & i + dec).Value

Print #1, Range("'Donn�es G�n�rales'!P" & i + dec + 5).Value
Print #1, Range("'Donn�es G�n�rales'!Q" & i + dec + 5).Value
Print #1, Range("'Donn�es G�n�rales'!R" & i + dec + 5).Value
Print #1, Range("'Donn�es G�n�rales'!S" & i + dec + 5).Value
Next i

Print #1, Range("'Donn�es G�n�rales'!R11").Value
Print #1, Range("'Donn�es G�n�rales'!R13").Value
Print #1, Range("'Donn�es G�n�rales'!R14").Value
Print #1, Range("'Donn�es G�n�rales'!S13").Value
Print #1, Range("'Donn�es G�n�rales'!S14").Value

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


'Caract�ristiques du bateau
Range("'Donn�es G�n�rales'!B3").Value = x(0)
Range("'Donn�es G�n�rales'!B4").Value = x(1)
Range("'Donn�es G�n�rales'!B5").Value = x(2)
Range("'Donn�es G�n�rales'!B8").Value = x(3)
Range("'Donn�es G�n�rales'!B10").Value = x(4)
Range("'Donn�es G�n�rales'!B11").Value = x(5)
Range("'Donn�es G�n�rales'!B12").Value = x(6)
Range("'Donn�es G�n�rales'!B13").Value = x(7)
 

'Fonction F1
Range("'Donn�es G�n�rales'!F6").Value = x(8)
Range("'Donn�es G�n�rales'!G6").Value = x(9)
Range("'Donn�es G�n�rales'!F9").Value = x(10)
Range("'Donn�es G�n�rales'!G9").Value = x(11)
Range("'Donn�es G�n�rales'!F12").Value = x(12)
Range("'Donn�es G�n�rales'!G12").Value = x(13)
Range("'Donn�es G�n�rales'!F15").Value = x(14)
Range("'Donn�es G�n�rales'!G15").Value = x(15)

'Fonction H1
Range("'Donn�es G�n�rales'!K6").Value = x(16)
Range("'Donn�es G�n�rales'!L6").Value = x(17)
Range("'Donn�es G�n�rales'!K9").Value = x(18)
Range("'Donn�es G�n�rales'!L9").Value = x(19)
Range("'Donn�es G�n�rales'!K12").Value = x(20)
Range("'Donn�es G�n�rales'!L12").Value = x(21)
Range("'Donn�es G�n�rales'!K15").Value = x(22)
Range("'Donn�es G�n�rales'!L15").Value = x(23)

'Noeuds Ai, y et z
dec = 3

For i = 0 To 1

'Range("'Donn�es G�n�rales'!C" & 2 * i + dec).Value = x(24 + 11 * i)
Range("'Donn�es G�n�rales'!P" & i + dec).Value = x(24 + 8 * i)
Range("'Donn�es G�n�rales'!Q" & i + dec).Value = x(25 + 8 * i)
Range("'Donn�es G�n�rales'!R" & i + dec).Value = x(26 + 8 * i)
Range("'Donn�es G�n�rales'!S" & i + dec).Value = x(27 + 8 * i)

Range("'Donn�es G�n�rales'!P" & i + dec + 5).Value = x(28 + 8 * i)
Range("'Donn�es G�n�rales'!Q" & i + dec + 5).Value = x(29 + 8 * i)
Range("'Donn�es G�n�rales'!R" & i + dec + 5).Value = x(30 + 8 * i)
Range("'Donn�es G�n�rales'!S" & i + dec + 5).Value = x(31 + 8 * i)



'Range("'Donn�es G�n�rales'!D" & 2 * i + dec).Value = x(25 + 11 * i)
'Range("'Donn�es G�n�rales'!E" & 2 * i + dec).Value = x(26 + 11 * i)
'Range("'Donn�es G�n�rales'!F" & 2 * i + dec).Value = x(27 + 11 * i)
'Range("'Donn�es G�n�rales'!G" & 2 * i + dec).Value = x(28 + 11 * i)
'Range("'Donn�es G�n�rales'!H" & 2 * i + dec).Value = x(29 + 11 * i)
'Range("'Donn�es G�n�rales'!I" & 2 * i + dec).Value = x(30 + 11 * i)
'Range("'Donn�es G�n�rales'!J" & 2 * i + dec).Value = x(31 + 11 * i)
'Range("'Donn�es G�n�rales'!K" & 2 * i + dec).Value = x(32 + 11 * i)
'Range("'Donn�es G�n�rales'!L" & 2 * i + dec).Value = x(33 + 11 * i)
'Range("'Donn�es G�n�rales'!M" & 2 * i + dec).Value = x(34 + 11 * i)

Next i

Range("'Donn�es G�n�rales'!R11").Value = x(40)
If x(41) = 0 Then x(41) = 0.01
If x(43) = 0 Then x(43) = 0.01
Range("'Donn�es G�n�rales'!R13").Value = x(41)
Range("'Donn�es G�n�rales'!R14").Value = x(42)
Range("'Donn�es G�n�rales'!S13").Value = x(43)
Range("'Donn�es G�n�rales'!S14").Value = x(44)

Close #1



End Sub
