Function HarfNotu(ortalama As Double) As String

    ortalama = ortalama / 8
    
    Select Case ortalama
        Case Is >= 90
            HarfNotu = "AA"
        Case Is >= 85
            HarfNotu = "BA"
        Case Is >= 80
            HarfNotu = "BB"
        Case Is >= 75
            HarfNotu = "CB"
        Case Is >= 70
            HarfNotu = "CC"
        Case Is >= 65
            HarfNotu = "DC"
        Case Is >= 60
            HarfNotu = "DD"
        Case Else
            HarfNotu = "FF"
    End Select
End Function

