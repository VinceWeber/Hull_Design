Attribute VB_Name = "Module8"
Option Explicit

Sub Calcul_greement()
Attribute Calcul_greement.VB_ProcData.VB_Invoke_Func = "n\n14"

' Calculs des efforts v�liques, aquisitions des donn�es sur la feuille
'Puis rotation uniquement autour du point B suivant l'axe z (gamma)
'(r�aliste seulement pour la GV)
'Il faudra prendre en compte la rotation du foc autour de l'�tai
'Puis rotation des points suivants g�te(alpha), et assi�tte (beta)
'
'D�clarations

Dim LHT, pos_mat, h_mat, lg_b�me_pcent, type_gr, Bt_dehors_pcent, A, b, H_b�me_pcent As Single
Dim Xa, Ya, Za, Xb, Yb, Zb, Xc, Yc, Zc As Single
Dim xD, yD, Zd, Xe, Ye, Ze, Xf, Yf, Zf As Single
Dim Xgv, Ygv, Zgv, Xfoc, Yfoc, Zfoc As Single
Dim Xcv, Ycv, Zcv, Sgv, Sfoc, Sv As Single
Dim Xcvr, Ycvr, Zcvr, Sgvr, Sfocr, Svr As Single
Dim alpha, beta, gamma, norme As Single
Dim Ar(3), Br(3), Cr(3), n(3) As Single
Dim xg, yg, zg, Mx, My, Mz As Single
Dim i As Integer
Dim v_vent, kz As Single

'Pi d�j� calcul� dans le module 2 !!!
Pi = 4 * Atn(1)
fb = Range("'Donn�es G�n�rales'!B13").Value


'Acquisition des donn�es

LHT = Range("'Gr��ment'!B2").Value
pos_mat = Range("'Gr��ment'!B3").Value
h_mat = Range("'Gr��ment'!B4").Value
lg_b�me_pcent = Range("'Gr��ment'!B5").Value
type_gr = Range("'Gr��ment'!B6").Value
Bt_dehors_pcent = Range("'Gr��ment'!B7").Value
A = Range("'Gr��ment'!B8").Value
b = Range("'Gr��ment'!B9").Value
H_b�me_pcent = Range("'Gr��ment'!B10").Value
v_vent = Range("'Gr��ment'!C14").Value

xg = Range("'Resultats'!G7").Value
yg = Range("'Resultats'!G8").Value
zg = Range("'Resultats'!G9").Value


'Pr�paration des calculs, calcul des coordonn�es des somm�ts de voiles

    Xa = LHT * pos_mat
    Ya = 0
    Za = fb + h_mat
    
    Xb = (pos_mat * LHT) * (1 - lg_b�me_pcent)
    Yb = 0
    Zb = H_b�me_pcent * h_mat + fb
    
    Xc = Xa
    Yc = 0
    Zc = Zb
                '********** Remarque
                ' Si on ne veut pas que le Foc monte en haut de l'�tai => il faut jouer sur les coordonn�es de D
    xD = Xa
    yD = 0
    Zd = type_gr * h_mat + fb
    
    Xe = (1 + Bt_dehors_pcent) * LHT
    Ye = 0
    Ze = fb
    
    Xf = LHT * (pos_mat + (1 - pos_mat) * b)
    Yf = 0
    Zf = fb + h_mat * A


'calculs des coordonn�es de centre de voilure
'Grand Voile :

Xgv = 1 / 3 * (Xa + Xb + Xc)
Ygv = 1 / 3 * (Ya + Yb + Yc)
Zgv = 1 / 3 * (Za + Zb + Zc)
Sgv = 1 / 2 * (Za - Zc) * (Xc - Xb)

'Foc
Xfoc = 1 / 3 * (xD + Xe + Xf)
Yfoc = 1 / 3 * (yD + Ye + Yf)
Zfoc = 1 / 3 * (Zd + Ze + Zf)

If (xD - Xf) = 0 Then
Sfoc = Abs(1 / 2 * ((Xe - Xf) ^ 2 + (Ze - Zf) ^ 2) ^ (1 / 2) * ((xD - Xf) ^ 2 + (Zd - Zf) ^ 2) ^ (1 / 2) * Sin(Pi / 2 - Atn((Ze - Zf) / (Xe - Xf))))
ElseIf Xe - Xf = 0 Then
Sfoc = Abs(1 / 2 * ((Xe - Xf) ^ 2 + (Ze - Zf) ^ 2) ^ (1 / 2) * ((xD - Xf) ^ 2 + (Zd - Zf) ^ 2) ^ (1 / 2) * Sin(Atn((Zd - Zf) / (xD - Xf)) - Pi / 2))
Else
Sfoc = Abs(1 / 2 * ((Xe - Xf) ^ 2 + (Ze - Zf) ^ 2) ^ (1 / 2) * ((xD - Xf) ^ 2 + (Zd - Zf) ^ 2) ^ (1 / 2) * Sin(Atn((Zd - Zf) / (xD - Xf)) - Atn((Ze - Zf) / (Xe - Xf))))
End If

'Centre de Voilure
Sv = Sfoc + Sgv
Xcv = 1 / Sv * (Sgv * Xgv + Sfoc * Xfoc)
Ycv = 1 / Sv * (Sgv * Ygv + Sfoc * Yfoc)
Zcv = 1 / Sv * (Sgv * Zgv + Sfoc * Zfoc)

            alpha = Val(Range("'gr��ment'!C15").Value) * Pi / 180
            beta = Val(Range("'gr��ment'!C16").Value) * Pi / 180
            gamma = Val(Range("'gr��ment'!C17").Value) * Pi / 180

'GV
Xcvr = Xcv
Ycvr = Ycv
Zcvr = Zcv



Ar(1) = Xa
Ar(2) = Ya
Ar(3) = Za
Br(1) = Xb
Br(2) = Yb
Br(3) = Zb
Cr(1) = Xc
Cr(2) = Yc
Cr(3) = Zc

'rotation de la bome appliqu�e uniquement au point B

Br(1) = Xa - (Xa - Br(1)) * Cos(gamma)
Br(2) = Sin(gamma) * Br(1)
Xcvr = Xa - (Xa - Xcvr) * Cos(gamma)
Ycvr = Sin(gamma) * Ycvr

'Rotation de gite et d'assiette
rotation_point Ar(1), Ar(2), Ar(3), alpha, beta
rotation_point Br(1), Br(2), Br(3), alpha, beta
rotation_point Cr(1), Cr(2), Cr(3), alpha, beta
rotation_point Xcvr, Ycvr, Zcvr, alpha, beta
Sv = Sv * Cos(alpha)


Range("'Gr��ment'!C24").Value = Xcvr
Range("'Gr��ment'!C25").Value = Ycvr
Range("'Gr��ment'!C26").Value = Zcvr

n(1) = (Cr(2) - Ar(2)) * (Br(3) - Ar(3)) - (Cr(3) - Ar(3)) * (Br(2) - Ar(2))
n(2) = (Cr(3) - Ar(3)) * (Br(1) - Ar(1)) - (Cr(1) - Ar(1)) * (Br(3) - Ar(3))
n(3) = (Cr(1) - Ar(1)) * (Br(2) - Ar(2)) - (Cr(2) - Ar(2)) * (Br(1) - Ar(1))

norme = (n(1) ^ 2 + n(2) ^ 2 + n(3) ^ 2) ^ (1 / 2)

For i = 1 To 3
n(i) = n(i) / norme
Next i


'***** Application d'une valeur de force v�lique sur le vecteur normal
' Force proportionelle � la surface

For i = 1 To 3
n(i) = n(i) * Sv / 10000 * (1 / 10 * v_vent ^ 2)
Next i
Range("'Gr��ment'!C19").Value = Sv / 10000 * (1 / 10 * v_vent ^ 2)
Range("'Gr��ment'!C20").Value = n(1)
Range("'Gr��ment'!C21").Value = n(2)
Range("'Gr��ment'!C22").Value = n(3)

'affichage sur la feuille r�sultats

For i = 0 To 20
If Range("'Resultats'!T" & i + 13).Value <> "" Then
kz = Range("'Resultats'!T" & i + 13).Value
Mx = Ycv * n(3) - (Zcv - kz) * n(2)
Range("'Resultats'!U" & i + 13).Value = Mx
Range("'Resultats'!V" & i + 13).Value = "=U" & i + 13 & "+F" & i + 13



End If
Next i



End Sub

Sub rotation_point(x, y, z, alpha, beta)
Dim xx, yy, zz As Single

'rotation dans le plan Oyz (g�te)
yy = Cos(alpha) * y + Sin(alpha) * z
zz = -Sin(alpha) * y + Cos(alpha) * z

'rotation dans le plan Oxz (assiette)

xx = Cos(beta) * x + Sin(beta) * zz
zz = -Sin(beta) * x + Cos(beta) * zz

x = xx
y = yy
z = zz

End Sub

Sub graphique()
Dim h_mat, prof_max, longueur, x, ymin, ymax, fb As Single

h_mat = Range("'Gr��ment'!B4").Value
prof_max = Range("'Donn�es g�n�rales'!B10").Value
longueur = Range("'Donn�es g�n�rales'!B3").Value
fb = Range("'Donn�es g�n�rales'!B13").Value

If longueur > h_mat + prof_max Then
    
      x = longueur * 1.25
      ymin = -prof_max * 1.25
      ymax = (longueur - prof_max) * 1.25
    
    Else
    
    x = h_mat * 1.25 + prof_max
    ymin = -prof_max
    ymax = h_mat * 1.25
    
    End If
        
  
    
    
    ActiveSheet.ChartObjects("Graphique " & 1).Activate
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
    
End Sub
