Function HarfNotu2(KRDortalama As Double) As String

    KRDortalama = ortalama / 23
    
    Select Case KRDortalama
        Case Is >= 90
            HarfNotu2 = "AA"
        Case Is >= 85
            HarfNotu2 = "BA"
        Case Is >= 80
            HarfNotu2 = "BB"
        Case Is >= 75
            HarfNotu2 = "CB"
        Case Is >= 70
            HarfNotu2 = "CC"
        Case Is >= 65
            HarfNotu2 = "DC"
        Case Is >= 60
            HarfNotu2 = "DD"
        Case Else
            HarfNotu2 = "FF"
    End Select
End Function

