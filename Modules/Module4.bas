Attribute VB_Name = "Module4"
Option Explicit
Sub export_flottaisons()
Dim s, i, compt As Integer
Dim xmax, xmin, bwl_max, yd1, yd2, T, sx, sx2, sy, sy2, sxy, den, A, b, aire, air1, air2 As Single
Dim col As String
Dim premier As Boolean

For i = 0 To nb_ang + 1
xmax = -500
premier = False
bwl_max = 0
T = 0
sx = 0
sx2 = 0
sy = 0
sxy = 0
den = 0
compt = 0
aire = 0
'air1 = 0
'air2 = 0
For s = 0 To nu
If p(i, s, 0, 0) <> -100 Then
        
        sx = sx + p(i, s, 0, 0)
        sx2 = sx2 + p(i, s, 0, 0) ^ 2
        sxy = sxy + p(i, s, 0, 0) * ((p(i, s, 1, 1) + p(i, s, 0, 1)) / 2)
        sy = sy + ((p(i, s, 1, 1) + p(i, s, 0, 1)) / 2)
        compt = compt + 1
        
        If s <> nu_min Then
        aire = aire + ((p(i, s, 1, 1) - p(i, s, 0, 1)) + (p(i, s + 1, 1, 1) - p(i, s + 1, 0, 1))) / 2 * (p(i, s + 1, 0, 0) - p(i, s, 0, 0))
        'air1 = air1 + (p(i, s, 1, 1) + p(i, s + 1, 1, 1)) * (p(i, s + 1, 0, 0) - p(i, s, 0, 0)) / 2
       ' air2 = air2 + (p(i, s, 0, 1) + p(i, s + 1, 0, 1)) * (p(i, s + 1, 0, 0) - p(i, s, 0, 0)) / 2
        End If
        
        If premier = False Then
        xmin = p(i, s, 0, 0)
        premier = True
        End If

       If p(i, s, 0, 0) > xmax Then
        xmax = p(i, s, 0, 0)
        End If

If p(i, s, 0, 2) <> -4500 And p(i, s, 0, 0) <> -100 Then
       yd1 = p(i, s, 0, 1)
       yd2 = p(i, s, 1, 1)
       
       If Abs(yd2 - yd1) > bwl_max Then bwl_max = Abs(yd2 - yd1)
       'If p(i, s, 0, 1) < yd1_max Then yd1_max = p(i, s, 0, 1)
       'If p(i, s, 1, 1) > yd2_max Then yd2_max = p(i, s, 1, 1)
        
End If

If Abs(p(i, s, 1, 2)) > T Then T = Abs(p(i, s, 1, 2))


End If
Next s

'Calcul des axes principaux
den = compt * sx2 - sx ^ 2
'b = (sy * sx2 - sxy * sx) / den
A = (compt * sxy - sy * sx) / den

result(i, 13) = xmax - xmin 'LWL
result(i, 14) = bwl_max 'BWL
result(i, 15) = T 'T
result(i, 16) = p(i, 0, 0, 3) 'Aire du maître bau
result(i, 17) = p(i, 0, 1, 3) 'position du maitre bau (non remis dans le plan sans assiette)
result(i, 18) = Atn(A) * 180 / Pi
'result(i, 19) = aire

'p(num_angle, s, 0, 1) => yd1
'p(num_angle, s, 1, 1) =>yd2

'Convert_colonne 2 * i + 1, col
'Range("'vide'!" & col & 1).Value = "xmin="
'Range("'vide'!" & col & 2).Value = "xmax="
'Range("'vide'!" & col & 3).Value = "long_flott"
'Convert_colonne 2 * (i + 1), col
'Range("'vide'!" & col & 1).Value = xmin
'Range("'vide'!" & col & 2).Value = xmax
'Range("'vide'!" & col & 3).Value = xmax - xmin

Next i
End Sub
