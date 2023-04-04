Attribute VB_Name = "Module3"


Option Explicit
Sub acquisition_polynomes()
Dim i, s As Integer
Dim col As String

'Acquisition des paramètres du bateau:
fb = Range("'Données Générales'!B13").Value



'Acquisition des noeuds de Bezier

For s = 0 To 5
For i = 0 To 10
    Convert_colonne i + 3, col
    Py(s, i) = Range("'Polynomes'!" & col & 3 * (s + 1)).Value
    Pz(s, i) = Range("'Polynomes'!" & col & 3 * (s + 1) + 1).Value
Next i
Next s
'Rem: au lieu de faire une approximation polynomiale des noeuds de beziers sur toute la longueur
' préferer une approx par 'des splines' ou même une surface de Bezier
'=> moins de risque d'oscillation
' dans la feuille polynomes : de C27 à M39 représente les coordonnées des noeuds de Bezier
' Le reste de la feuille correspond au calcul des coefficients du polynôme d'approximation
' dont le résultat se situe dans la plage C3 M19
'Ce changement affectera seulement les Sub aquisition polynome et les fonctions Ay et Az
End Sub
Sub creation_couples()
Dim long_max, x, y, z, T As Single
Dim i, n, s As Integer
Dim col As String

long_max = Range("'P(F1)'!M18").Value

For i = 0 To nu
x = i * long_max / nu
    For n = 0 To 5
        Bezier(i, n, 0) = x
        Bezier(i, n, 1) = Ay(x, n)
        Bezier(i, n, 2) = Az(x, n)
    Next n
Next i

For s = 0 To nu
For i = 0 To nv
T = i / nv
x = s * long_max / nu
points3(s, i, 0) = x
y = Bezier(s, 0, 1) * (1 - T) ^ 5 + 5 * Bezier(s, 1, 1) * (1 - T) ^ 4 * T + 10 * Bezier(s, 2, 1) * (1 - T) ^ 3 * T ^ 2 + 10 * Bezier(s, 3, 1) * (1 - T) ^ 2 * T ^ 3 + 5 * Bezier(s, 4, 1) * (1 - T) * T ^ 4 + Bezier(s, 5, 1) * T ^ 5
z = Bezier(s, 0, 2) * (1 - T) ^ 5 + 5 * Bezier(s, 1, 2) * (1 - T) ^ 4 * T + 10 * Bezier(s, 2, 2) * (1 - T) ^ 3 * T ^ 2 + 10 * Bezier(s, 3, 2) * (1 - T) ^ 2 * T ^ 3 + 5 * Bezier(s, 4, 2) * (1 - T) * T ^ 4 + Bezier(s, 5, 2) * T ^ 5

points3(s, nv + i, 1) = y
points3(s, nv - i, 1) = -y
points3(s, nv + i, 2) = z
points3(s, nv - i, 2) = z

Next i
Next s


End Sub

Public Function Ay(x, i)
' x= position du couple
' i= n° du noeud ( de 0 à 5)
Dim n As Integer
Dim Somme As Single
    For n = 0 To 10
        Somme = Somme + Py(i, n) * x ^ n
    Next n
    
Ay = Somme
End Function
Public Function Az(x, i)
' x= position du couple
' i= n° du noeud ( de 0 à 5)
Dim n As Integer
Dim Somme As Single
    For n = 0 To 10
        Somme = Somme + Pz(i, n) * x ^ n
    Next n
    
Az = Somme
End Function

