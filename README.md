Option Explicit
Sub Code_Canopy()

Dim b As Double
Dim c As Double

b = 0
c = 0
    

    Do While b < 1439
        If Range("b2").Offset(b, 0).Value < 30 Then
        Range("d2").Offset(b, 0).Value = 0
        Range("e2").Offset(b, 0).Value = Range("b2").Offset(b, 0).Value
        Else
        Range("d2").Offset(b, 0).Value = Range("b2").Offset(b, 0).Value - 30
        Range("e2").Offset(b, 0).Value = 30
        End If
        
        If Range("c2").Offset(b, 0).Value = "NULL" Then
        Range("c2").Offset(b, 0).Value = (Range("c2").Offset(b - 1, 0).Value + Range("c2").Offset(b + 1, 0).Value) / 2
        End If
        
        Range("a2").Offset(b, 0).Value = Range("a2").Offset(b, 0).Value + 1 / 24 / 60
        
        If Range("a2").Offset(b, 0).Value > 0.0409 And Range("a2").Offset(b, 0) < 0.0415 Then
        Range("f2").Offset(b, 0).Value = 1
        End If
        
        b = b + 1
        
    Loop
    
    b = 0
    c = 0
    
    Do While b < 1437
        c = c + Range("f2").Offset(1437 - b, 0).Value
        Range("g2").Offset(1437 - b - 1, 0).Value = c
        Range("a2").Offset(1437 - b - 1, 0).Value = c / 24 + Range("a2").Offset(1437 - b - 1, 0).Value
        b = b + 1
        
    Loop
 
    Range("h2").Value = Application.WorksheetFunction.Average(Range("d2:d1437")) * 24 * 0.3
    Range("d1").Value = "Diesel Load"
    Range("e1").Value = "Solar Load"
    Range("h1").Value = "Total Fuel Consumed"

End Sub

















