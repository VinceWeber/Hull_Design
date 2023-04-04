Attribute VB_Name = "Module2"
Global points4(400, 400, 2), Pi, precision_v, precision_xg, angle_max, nb_ang As Single
Global nb_points, nb_sections, nu_max, nv_max, nu_min, nv_min, nu, nv As Integer
Global Bezier(400, 5, 2), Py(5, 10), Pz(5, 10), points3(400, 600, 2), p(50, 400, 1, 3) As Single
Global result(50, 24), result_fixe(1, 6) As Single
Global cgx, cgy, cgz As Single
Global export_aires As Boolean

Global fb As Single

Global abandon As Boolean
Option Explicit


Sub Calcul(Calc_ass, masse)
Attribute Calcul.VB_ProcData.VB_Invoke_Func = " \n14"

Dim i, s, itt, itt2 As Integer
Dim V, enf, enfoncement(50), angle(50), beta(50), masse_desiree As Single
Dim xg, zg, yg, alpha, ro, v_dep, xcc, zcc, ycc, emax, emin As Single
Dim A, b, Sm, tab_arr, beta_min, beta_max, aa, bb, Crx, Cry, var_sm, err_vol As Single
Dim fin, max, carene, assiette, sortie As Boolean
Dim c(3), z1Ka, Zk, ycc1, xcc2 As Single
Dim col As String


'Pi = 4 * Atn(1)
nu = nu_max
nv = nv_max

'c(2) = Val(Range("'Gréément'!C34").Value)
c(2) = 0

'acquisition_polynomes
creation_couples

emax = fb
emin = -10 * Range("'Données Générales'!B10").Value

beta_min = -10 * Pi / 180
beta_max = 10 * Pi / 180


'masse volumique
        ro = 1.025

'Création des angles de gîte à calculer
        
        For i = 0 To nb_ang
        angle(i) = i * angle_max / nb_ang * Pi / 180
        beta(i) = 0
        Next i

'premiers calculs volumes



'coque entière=> Sotie : V, xg,zg,yg coord du centre de gravité
calcul_volume fb, 0, xg, zg, yg, V, Sm, tab_arr, 0, nb_ang + 1

result_fixe(0, 0) = V
result_fixe(0, 1) = Sm
result_fixe(0, 2) = xg
result_fixe(0, 3) = yg
result_fixe(0, 4) = zg

'carène à plat

calcul_volume 0, 0, xg, zg, yg, v_dep, Sm, tab_arr, 0, 0

If Form2.CheckBox4.Value = False Then
v_dep = masse / ro * 1000
End If

result_fixe(1, 0) = v_dep
result_fixe(1, 1) = Sm
result_fixe(1, 2) = xg
result_fixe(1, 3) = yg
result_fixe(1, 4) = zg
result_fixe(1, 5) = v_dep * ro / 1000

'initialisation du centre de gravité au niveau du centre de carène

If Form2.CheckBox3.Value = True Then
cgx = xg
cgy = yg
cgz = zg
Else
'cgx = Val(Form2.TextBox9.Text)
cgx = xg 'Placement du centre de gravité à la meme position x du centre de carene
cgy = Val(Form2.TextBox10.Text)
cgz = Val(Form2.TextBox11.Text)
End If

'carène avec angle et ittérations sur l'enfoncement (calcul approx)

nu = nu_min
nv = nv_min
creation_couples


For i = 0 To nb_ang

Form1.Lbl_angle = i & " / " & Form2.TextBox2.Text

aa = beta_min
bb = beta_max
enf = 0

itt2 = 0
assiette = False
sortie = False

Do While assiette = False And abandon = False And sortie = False
DoEvents
A = emin
b = emax
V = 0
itt = 0
calcul_volume enf, angle(i), xcc, zcc, ycc, V, Sm, tab_arr, beta(i), i

Do While Abs(V - v_dep) > precision_v And itt < 200 And abandon = False
DoEvents

        If V - v_dep > 0 Then
            b = enf
            enf = (enf + A) / 2
        Else
            A = enf
            enf = (enf + b) / 2
        End If
        
calcul_volume enf, angle(i), xcc, ycc, zcc, V, Sm, tab_arr, beta(i), i

itt = itt + 1
Form1.lbl_enf = itt

Loop
    If Calc_ass = False Then
    sortie = True
    Else
            If Abs((xcc - cgx) * v_dep * 9.81 / 1000 / 100 - c(2)) < precision_xg Then
           
                assiette = True
            Else

'xcc2 = Cos(beta(i)) * (xcc - cgx) + Sin(beta(i)) * (-Sin(angle(i)) * (ycc - cgy) + Cos(angle(i)) * (zcc - cgz)) + cgx
                
                If (xcc - cgx) * v_dep - c(2) > 0 Then
                
                'If (xcc2 - cgx) * v_dep * 9.81 / 100 - c(2) > 0 Then
                bb = beta(i)
                beta(i) = (beta(i) + aa) / 2
                Else
                aa = beta(i)
                beta(i) = (beta(i) + bb) / 2
                End If
             itt2 = itt2 + 1
            End If
    
    Form1.Lbl_assiette = itt2
    End If
                        
Loop
enfoncement(i) = enf


Next i


'carène avec angle et ittérations sur l'enfoncement (calcul final)

nu = nu_max
nv = nv_max
creation_couples


'effacement de la feuille "Aires"
                Sheets("Aires").Select
                Rows("6:150").ClearContents
'Validation de l'exportation des aires
        If Form2.CheckBox5.Value = True Then
        export_aires = True
        End If

If abandon = False Then

For i = 0 To nb_ang

calcul_volume enfoncement(i), angle(i), xcc, zcc, ycc, V, Sm, tab_arr, beta(i), i

'Convert_colonne i + 2, col


metacentre xcc, ycc, zcc, angle(i), beta(i), z1Ka, Zk, ycc1, xcc2

'Crx = (ycc - cgz * Sin(angle(i))) * V_dep * 9.81 / 1000
'Cry = (xcc - cgx) * V_dep * 9.81 / 1000

Crx = v_dep * 9.81 / 1000 * (ycc1 - cgy)
Cry = v_dep * 9.81 / 1000 * (xcc2 - cgx)


err_vol = (V - v_dep) / v_dep * 100

If i = 0 Then
var_sm = 0
Else
var_sm = (Sm - result(0, 5)) / result(0, 5) * 100
End If

result(i, 0) = angle(i) * 180 / Pi
result(i, 1) = beta(i) * 180 / Pi
result(i, 2) = xcc
result(i, 3) = ycc
result(i, 4) = zcc
result(i, 5) = Sm
result(i, 6) = var_sm
result(i, 7) = Crx
result(i, 8) = Cry
result(i, 9) = tab_arr
result(i, 10) = enfoncement(i)
result(i, 11) = V
result(i, 12) = err_vol
result(i, 20) = Zk

Next i

Range("'Données Générales'!B18").Value = enfoncement(0)

End If
End Sub

Sub metacentre(xcc, ycc, zcc, alpha, beta, z1Ka, Zk, ycc1, xcc2)
Dim xx, yy, zz, y1, z1 As Single
Dim zcc1 As Single
'Translation au Cg
xcc2 = xcc - cgx
ycc1 = ycc - cgy
zcc1 = zcc - cgz
'rotation des ycc,zcc dans le plan Oyz (gite)
yy = Cos(alpha) * ycc1 + Sin(alpha) * zcc1
zz = -Sin(alpha) * ycc1 + Cos(alpha) * zcc1
'rotation de l'axe x dans le plan Oxz (assiette)
xx = Cos(beta) * xcc2 + Sin(beta) * zz
'translation inverse
xcc2 = xx + cgx
ycc1 = yy + cgy
zcc1 = zz + cgz

'If beta <> 0 Then
'z1Ka = zcc1 - Tan(-(Pi / 2 - beta)) * ycc1
'Else
'z1Ka = cgz * Cos(alpha)
'End If

If alpha <> 0 Then
Zk = zcc - Tan(-(Pi / 2 - alpha)) * ycc
Else
Zk = 0
End If

End Sub
