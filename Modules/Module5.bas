Attribute VB_Name = "Module5"
Option Explicit
Sub calcul_volume(enf, alpha, xg, zgg, ygg, V, Sm, tab_arr, beta, num_angle)
Dim premier, dernier As Boolean
Dim yD, xD, y1, y2, z1, z2, x1, x2, A(200), yg(200), zg(200), zmini(200) As Single
Dim Za, Zb, x, y, z, surf(200) As Single
Dim s, i As Integer
Dim col As String

premier = False
dernier = False

'Rotation d'angle alpha
rotation alpha, beta
'calcul de l'aire & centre de Gravité
For s = 0 To nu
A(s) = 0
yg(s) = 0
zg(s) = 0
surf(s) = 0
Next s

For s = 0 To nu
zmini(s) = 10000000000#

For i = 1 To 2 * nv
z2 = points4(s, i, 2)
z1 = points4(s, i - 1, 2)
y2 = points4(s, i, 1)
y1 = points4(s, i - 1, 1)
Za = z1 - enf
Zb = z2 - enf
p(num_angle, s, 0, 0) = points4(s, 0, 0)
                    'recherche du z minimum
                    If z2 < zmini(s) Then
                    zmini(s) = z2
                    p(num_angle, s, 1, 2) = Zb
                    End If
    
    If z2 < enf Then
        If premier = False Then
             yD = (-Za) * (y2 - y1) / (Zb - Za) + y1
            p(num_angle, s, 0, 1) = yD
                If yD > y2 Then
                  A(s) = -(Zb) * (y2 - yD) / 2
                  yg(s) = -Zb * (y2 - yD) / 2 * (yD + y2) / 2
                  zg(s) = -Zb * (y2 - yD) / 2 * z2 / 2
                Else
                  A(s) = (Zb) * (y2 - yD) / 2
                  yg(s) = Zb * (y2 - yD) / 2 * (yD + y2) / 2
                  zg(s) = Zb * (y2 - yD) / 2 * z2 / 2
                End If
                  surf(s) = ((yD - y2) ^ 2 + (Za - Zb) ^ 2) ^ (1 / 2)
            premier = True
        
        Else
        
            If y1 > y2 Then
                   A(s) = A(s) - (Za + Zb) * (y2 - y1) / 2
                   yg(s) = yg(s) - (Zb + Za) * (y2 - y1) / 2 * (y1 + y2) / 2
                   zg(s) = zg(s) - (Za + Zb) * (y2 - y1) / 2 * (z1 + z2) / 2
            Else
                   A(s) = A(s) + (Za + Zb) * (y2 - y1) / 2
                   yg(s) = yg(s) + (Zb + Za) * (y2 - y1) / 2 * (y1 + y2) / 2
                   zg(s) = zg(s) + (Za + Zb) * (y2 - y1) / 2 * (z1 + z2) / 2
            End If
                   surf(s) = surf(s) + ((y1 - y2) ^ 2 + (Za - Zb) ^ 2) ^ (1 / 2)
        End If
    Else
        If premier = True And dernier = False Then
            yD = (-Za) * (y2 - y1) / (Zb - Za) + y1
            p(num_angle, s, 1, 1) = yD
           If y1 > yD Then
            A(s) = A(s) - (Za) * (yD - y1) / 2
            yg(s) = yg(s) - Za * (yD - y1) / 2 * (y1 + yD) / 2
            zg(s) = zg(s) - Za * (yD - y1) / 2 * z1 / 2
           Else
            A(s) = A(s) + (Za) * (yD - y1) / 2
            yg(s) = yg(s) + Za * (yD - y1) / 2 * (y1 + yD) / 2
            zg(s) = zg(s) + Za * (yD - y1) / 2 * z1 / 2
           End If
           surf(s) = surf(s) + ((yD - y1) ^ 2 + (Za - Zb) ^ 2) ^ (1 / 2)
            dernier = True
        End If
    End If
Next i

'**************************
'Cas des angles importants => franc bord submergé

If points4(s, 2 * nv, 2) < enf And points4(s, 0, 2) > enf Then

z2 = points4(s, 0, 2)
z1 = points4(s, 2 * nv, 2)
y2 = points4(s, 0, 1)
y1 = points4(s, 2 * nv, 1)
Za = z1 - enf
Zb = z2 - enf

yD = (-Za) * (y2 - y1) / (Zb - Za) + y1
p(num_angle, s, 1, 1) = yD
p(num_angle, s, 0, 2) = -4500
      If alpha < Pi / 2 Then
        A(s) = A(s) - (Za) * (yD - y1) / 2
        yg(s) = yg(s) - Za * (yD - y1) / 2 * (y1 + yD) / 2
        zg(s) = zg(s) - Za * (yD - y1) / 2 * z1 / 2
      Else
        A(s) = A(s) + (Za) * (yD - y1) / 2
        yg(s) = yg(s) + Za * (yD - y1) / 2 * (y1 + yD) / 2
        zg(s) = zg(s) + Za * (yD - y1) / 2 * z1 / 2
      End If
        surf(s) = surf(s) + (yD ^ 2 + y1 ^ 2) ^ (1 / 2)
End If


'****************************

If zmini(s) < enf Then
A(s) = -A(s)

yg(s) = -yg(s) / A(s)
zg(s) = (-zg(s) / A(s) + enf) / 2

Else
A(s) = 0
yg(s) = 0
zg(s) = 0
surf(s) = 0
p(num_angle, s, 0, 0) = -100
p(num_angle, s, 0, 1) = 0
p(num_angle, s, 1, 1) = 0
End If

If export_aires = True Then
courbes_aires s, num_angle, A(s), alpha
End If

premier = False
dernier = False
Next s

V = 0
ygg = 0
zgg = 0
xg = 0
Sm = 0
p(num_angle, 0, 0, 3) = A(0)
For s = 1 To nu
z1 = zmini(s - 1)
z2 = zmini(s)
x1 = points4(s - 1, 0, 0)
x2 = points4(s, 0, 0)

'rech de l'aire maxi

If A(s) > p(num_angle, 0, 0, 3) Then
p(num_angle, 0, 0, 3) = A(s)
p(num_angle, 0, 1, 3) = x2
End If


If zmini(s - 1) > enf And zmini(s) < enf Then
        xD = (enf - z1) * (x2 - x1) / (z2 - z1) + x1
        V = V + A(s) * (x2 - xD) / 2
        xg = xg + A(s) * (x2 - xD) * (x2 + xD) / 4
        ygg = ygg + A(s) * yg(s) * (x2 - xD) / 2
        zgg = zgg + A(s) * zg(s) * (x2 - xD) / 2
        Sm = Sm + surf(s) * (x2 - xD) / 2
ElseIf zmini(s - 1) < enf And zmini(s) > enf Then
        xD = (enf - z1) * (x2 - x1) / (z2 - z1) + x1
        V = V + A(s - 1) * (xD - x1) / 2
        xg = xg + A(s - 1) * (xD - x1) * (xD + x1) / 4
        ygg = ygg + A(s - 1) * yg(s) * (xD - x1) / 2
        zgg = zgg + A(s - 1) * zg(s) * (xD - x1) / 2
        Sm = Sm + surf(s - 1) * (xD - x1) / 2
Else
        V = V + (A(s) + A(s - 1)) * (x2 - x1) / 2
        xg = xg + (A(s) + A(s - 1)) * (x2 - x1) * (x1 + x2) / 4
        ygg = ygg + (A(s) + A(s - 1)) * yg(s) * (x2 - x1) / 2
        zgg = zgg + (A(s) + A(s - 1)) * zg(s) * (x2 - x1) / 2
        Sm = Sm + (surf(s) + surf(s - 1)) * (x2 - x1) / 2
End If

Next s
premier = False

If V <> 0 Then
xg = xg / V
ygg = ygg / V
zgg = zgg / V
End If

tab_arr = A(0)

'Rotation inverse (retour au repère d'origine)
x = xg - cgx
y = ygg - cgy
z = zgg - cgz

xg = Cos(beta) * x - Sin(beta) * z
zgg = Sin(beta) * x + Cos(beta) * z

z = zgg

ygg = Cos(alpha) * y - Sin(alpha) * z
zgg = Sin(alpha) * y + Cos(alpha) * z

xg = xg + cgx
ygg = ygg + cgy
zgg = zgg + cgz

End Sub
Sub rotation(alpha, beta)
Dim s, i As Integer
Dim x, y, z As Single
Dim xx, yy, zz As Single
Dim col As String

For s = 0 To nu
For i = 0 To 2 * nv

'Translation dans le repère lié au centre de gravité
x = points3(s, 0, 0) - cgx
y = points3(s, i, 1) - cgy
z = points3(s, i, 2) - cgz


'rotation dans le plan Oyz (gîte)
yy = Cos(alpha) * y + Sin(alpha) * z
zz = -Sin(alpha) * y + Cos(alpha) * z

'rotation dans le plan Oxz (assiette)

xx = Cos(beta) * x + Sin(beta) * zz
zz = -Sin(beta) * x + Cos(beta) * zz

'Translation inverse
points4(s, 0, 0) = xx + cgx
points4(s, i, 1) = yy + cgy
points4(s, i, 2) = zz + cgz

Next i
Next s

End Sub
