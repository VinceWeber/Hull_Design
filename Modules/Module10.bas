Attribute VB_Name = "Module10"
Option Explicit

Sub surface_derive()

Dim n, i, pas, compteur(1) As Integer
Dim A(13), ph(10), x(1), lg, h(1), xD(4), xS(4), yS(4), yD(3) As Double
Dim g(1, 13) As Single
Dim col As String

Dim premier_A3, premier_A7 As Boolean

pas = 1000
premier_A3 = False
premier_A7 = False

For i = 0 To 5
    A(i) = 0
Next i

lg = Range("'Données Générales'!B9").Value
xD(1) = Range("'Plan derive'!C8").Value
xD(2) = Range("'Plan derive'!C9").Value
xD(3) = Range("'PLan derive'!C10").Value
xD(4) = Range("'Plan derive'!C11").Value
yD(1) = Range("'Plan derive'!B13").Value
yD(2) = Range("'Plan derive'!B14").Value
yD(3) = Range("'Plan derive'!B15").Value

xS(1) = Range("'Plan derive'!E8").Value
xS(2) = Range("'Plan derive'!E9").Value
xS(3) = Range("'PLan derive'!E10").Value
xS(4) = Range("'Plan derive'!E11").Value
yS(1) = Range("'Plan derive'!D13").Value
yS(2) = Range("'Plan derive'!D14").Value
yS(3) = Range("'Plan derive'!D15").Value
yS(4) = Range("'Plan derive'!D15").Value

For n = 0 To 10
    Convert_colonne n + 3, col
    ph(n) = Val(Range("'P(H1)'!" & col & 9).Value)
Next n

For i = 0 To pas

    x(0) = i * lg / pas
    x(1) = (i + 1) * lg / pas
    
    h(0) = 0
    h(1) = 0
    
    For n = 0 To 10
        h(0) = h(0) + ph(n) * x(0) ^ n
        h(1) = h(1) + ph(n) * x(1) ^ n
    Next n
    
    If h(1) < 0 And h(0) < 0 Then
        
        A(0) = A(0) + (h(0) + h(1)) / 2 * lg / pas
        g(0, 0) = g(0, 0) + (x(0) + x(1)) / 2 * (h(0) + h(1)) / 2 * lg / pas
        g(1, 0) = g(1, 0) + (h(0) + h(1)) / 2 * (h(0) + h(1)) / 2 * lg / pas
        
        If x(0) > xD(1) And x(0) < xD(2) Then
            If premier_A3 = False Then
                premier_A3 = True
                A(3) = A(3) + (h(0)) / 2 * lg / pas
                g(0, 3) = g(0, 3) + x(0) * h(0) / 2 * lg / pas
                g(1, 3) = g(1, 3) + h(0) * h(0) / 2 * lg / pas
            Else
                A(3) = A(3) + (h(0) + h(1)) / 2 * lg / pas
                g(0, 3) = g(0, 3) + (x(0) + x(1)) / 2 * (h(0) + h(1)) / 2 * lg / pas
                g(1, 3) = g(1, 3) + (h(0) + h(1)) / 2 * (h(0) + h(1)) / 2 * lg / pas
            End If
       End If
       
       If x(0) > xS(1) And x(0) < xS(2) Then
            If premier_A7 = False Then
                premier_A7 = True
                A(7) = A(7) + (h(0)) / 2 * lg / pas
                g(0, 7) = g(0, 7) + x(0) * h(0) / 2 * lg / pas
                g(1, 7) = g(1, 7) + h(0) * h(0) / 2 * lg / pas
            Else
                A(7) = A(7) + (h(0) + h(1)) / 2 * lg / pas
                g(0, 7) = g(0, 7) + (x(0) + x(1)) / 2 * (h(0) + h(1)) / 2 * lg / pas
                g(1, 7) = g(1, 7) + (h(0) + h(1)) / 2 * (h(0) + h(1)) / 2 * lg / pas
            End If
       End If
       
    ElseIf h(1) < 0 And h(0) > 0 Then
    
        A(0) = A(0) + h(1) / 2 * lg / pas
        g(0, 0) = g(0, 0) + x(1) * h(1) / 2 * lg / pas
        g(1, 0) = g(1, 0) + h(1) * h(1) / 2 * lg / pas
    
    ElseIf h(1) > 0 And h(0) < 0 Then
    
        A(0) = A(0) + h(0) / 2 * lg / pas
        g(0, 0) = g(0, 0) + x(0) * h(0) / 2 * lg / pas
        g(1, 0) = g(1, 0) + h(0) * h(0) / 2 * lg / pas
    
    End If
    
    
Next i

g(0, 0) = g(0, 0) / A(0)
g(1, 0) = g(1, 0) / (2 * A(0))
g(0, 3) = g(0, 3) / A(3)
g(1, 3) = g(1, 3) / (2 * A(3))
g(0, 7) = g(0, 7) / A(7)
g(1, 7) = g(1, 7) / (2 * A(7))

'Range("'Plan Derive'!B23").Value = g(0, 0)
'Range("'Plan Derive'!B24").Value = g(1, 0)
'Range("'Plan Derive'!B25").Value = g(0, 3)
'Range("'Plan Derive'!B26").Value = g(1, 3)

'A(0) = A(0) * lg / pas
'A(3) = -A(3)





'Range("'Plan Derive'!C18").Value = A(0)
A(1) = -1 / 2 * (xD(1) - xD(4)) * (yD(1) - yD(3))
g(0, 1) = 1 / 3 * (2 * xD(1) + xD(4))
g(1, 1) = 1 / 3 * (2 * yD(3) + yD(1))
'Range("'Plan Derive'!C19").Value = A(1)
A(2) = (xD(2) - xD(1)) * yD(3)
g(0, 2) = 1 / 2 * (xD(1) + xD(2))
g(1, 2) = 1 / 2 * yD(3)
'Range("'Plan Derive'!C20").Value = A(2)
'Range("'Plan Derive'!C21").Value = A(3)
A(4) = -1 / 2 * (xD(2) - xD(3)) * (yD(2) - yD(3))
g(0, 4) = 1 / 3 * (2 * xD(2) + xD(3))
g(1, 4) = 1 / 3 * (2 * yD(3) + yD(2))
'Range("'Plan Derive'!C22").Value = A(4)
A(5) = -1 / 2 * (xS(1) - xS(4)) * (yS(1) - yS(3))
g(0, 5) = 1 / 3 * (2 * xS(1) + xS(4))
g(1, 5) = 1 / 3 * (2 * yS(3) + yS(1))

A(6) = (xS(2) - xS(1)) * yS(4)
g(0, 6) = 1 / 2 * (xS(1) + xS(2))
g(1, 6) = 1 / 2 * yS(4)

A(8) = -1 / 2 * (xS(2) - xS(3)) * (yS(2) - yS(3))
g(0, 8) = 1 / 3 * (2 * xS(2) + xS(3))
g(1, 8) = 1 / 3 * (2 * yS(3) + yS(2))

A(9) = (xS(2) - xS(3)) * (yS(3) - yS(4))
g(0, 9) = 1 / 2 * (xS(2) + xS(3))
g(1, 9) = 1 / 2 * (yS(3) + yS(4))

A(10) = -1 / 2 * (xS(3) - xS(4)) * (yS(3) - yS(4))
g(0, 10) = 1 / 3 * (2 * xS(3) + xS(4))
g(1, 10) = 1 / 3 * (2 * yS(4) + yS(3))

'Quille
A(11) = A(2) + A(1) - A(3) - A(4)
g(0, 11) = 1 / A(11) * (g(0, 2) * A(2) + g(0, 1) * A(1) - g(0, 3) * A(3) - g(0, 4) * A(4))
g(1, 11) = 1 / A(11) * (g(1, 2) * A(2) + g(1, 1) * A(1) - g(1, 3) * A(3) - g(1, 4) * A(4))

'Safran
A(12) = A(6) + A(5) - A(7) - A(8) - A(9) - A(10)
g(0, 12) = 1 / A(12) * (g(0, 6) * A(6) + g(0, 5) * A(5) - g(0, 7) * A(7) - g(0, 8) * A(8) - g(0, 9) * A(9) - g(0, 10) * A(10))
g(1, 12) = 1 / A(12) * (g(1, 6) * A(6) + g(1, 6) * A(6) - g(1, 7) * A(7) - g(1, 8) * A(8) - g(1, 9) * A(9) - g(1, 10) * A(10))

'Global
A(13) = A(0) + A(12) + A(11)
g(0, 13) = 1 / A(13) * (g(0, 0) * A(0) + g(0, 12) * A(12) + g(0, 11) * A(11))
g(1, 13) = 1 / A(13) * (g(1, 0) * A(0) + g(1, 12) * A(12) + g(1, 11) * A(11))

'Range("'Plan Derive'!B23").Value = A(5)

Range("'Plan Derive'!B18").Value = -A(13)
Range("'Plan Derive'!B21").Value = g(0, 13)
Range("'Plan Derive'!B22").Value = g(1, 13)

Range("'Plan Derive'!E18").Value = -A(11)
Range("'Plan Derive'!E19").Value = -A(12)

End Sub
Public Function Hd()
' x= position du couple
' i= n° du noeud ( de 0 à 5)

End Function

Sub graphique_derive()

Dim h_mat, prof_max, longueur, x, ymin, ymax, fb As Single

h_mat = Range("'Gréément'!B4").Value
prof_max = Range("'Plan derive'!B4").Value
longueur = Range("'Données générales'!B3").Value
fb = Range("'Données générales'!B13").Value

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
