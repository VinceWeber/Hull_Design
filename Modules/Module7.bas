Attribute VB_Name = "Module7"
Option Explicit

Sub Lancer_calcul()
Attribute Lancer_calcul.VB_Description = "Macro enregistrée le 28/02/2004 par Vincent"
Attribute Lancer_calcul.VB_ProcData.VB_Invoke_Func = "t\n14"

Form2.Show

End Sub

Sub afficher_resultats()

Dim i, decalage As Integer

decalage = 13
Sheets("Resultats").Select
Rows(decalage + 1 & ":100").ClearContents

Range("'resultats'!D4").Value = result_fixe(0, 0)
Range("'resultats'!D5").Value = result_fixe(0, 1)
Range("'resultats'!G4").Value = result_fixe(0, 2)
Range("'resultats'!G5").Value = result_fixe(0, 3)
Range("'resultats'!G6").Value = result_fixe(0, 4)

Range("'resultats'!D7").Value = result_fixe(1, 0)
Range("'resultats'!D9").Value = result_fixe(1, 1)
Range("'resultats'!G7").Value = result_fixe(1, 2)
Range("'resultats'!G8").Value = result_fixe(1, 3)
Range("'resultats'!G9").Value = result_fixe(1, 4)
Range("'resultats'!D8").Value = result_fixe(1, 5)

' Coefficient prismatique
Range("'resultats'!N4") = result_fixe(1, 0) * 1.025 / (result(0, 13) * result(0, 16))
'Coefficient de bloc
Range("'resultats'!N5") = result_fixe(1, 0) * 1.025 / (result(0, 13) * result(0, 14) * result(0, 15))

Range("'resultats'!N7").Value = result(nb_ang + 1, 13)
Range("'resultats'!N8").Value = result(nb_ang + 1, 14)
' nu et nv

Range("'resultats'!R4").Value = Val(F_chargmt.TextBox5.Value)
Range("'resultats'!R5").Value = Val(F_chargmt.TextBox6.Value)




If Form2.CheckBox3.Value = False Then
'Range("'resultats'!L4").Value = Val(Form2.TextBox9.Text)
'Range("'resultats'!L5").Value = Val(Form2.TextBox10.Text)
'Range("'resultats'!L6").Value = Val(Form2.TextBox11.Text)
Range("'resultats'!L4").Value = cgx
Range("'resultats'!L5").Value = cgy
Range("'resultats'!L6").Value = cgz
Range("'resultats'!I7").Value = ""
Range("'resultats'!I8").Value = ""
Else
Range("'resultats'!I7").Value = "Calcul au centre de carène"
Range("'resultats'!I8").Value = "à gîte nulle"
Range("'resultats'!L4").Value = ""
Range("'resultats'!L5").Value = ""
Range("'resultats'!L6").Value = ""
End If


For i = 0 To nb_ang

Range("'resultats'!A" & i + decalage) = result(i, 0)
Range("'resultats'!B" & i + decalage) = result(i, 1)
Range("'resultats'!C" & i + decalage) = result(i, 2)
Range("'resultats'!D" & i + decalage) = result(i, 5)
Range("'resultats'!E" & i + decalage) = result(i, 6)
Range("'resultats'!F" & i + decalage) = result(i, 7)
Range("'resultats'!G" & i + decalage) = result(i, 9)
Range("'resultats'!H" & i + decalage) = result(i, 13)
Range("'resultats'!I" & i + decalage) = result(i, 14)
Range("'resultats'!J" & i + decalage) = result(i, 15)
Range("'resultats'!K" & i + decalage) = result(i, 16)
Range("'resultats'!L" & i + decalage) = result(i, 17)
Range("'resultats'!M" & i + decalage) = result(i, 18)
Range("'resultats'!N" & i + decalage) = result(i, 3)
Range("'resultats'!O" & i + decalage) = result(i, 4)
Range("'resultats'!P" & i + decalage) = result(i, 8)
Range("'resultats'!Q" & i + decalage) = result(i, 10)
Range("'resultats'!R" & i + decalage) = result(i, 11)
Range("'resultats'!S" & i + decalage) = result(i, 12)
Range("'resultats'!T" & i + decalage) = result(i, 20)

Next i


End Sub
