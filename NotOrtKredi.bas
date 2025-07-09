Function NotOrtKredi() As Double
    Dim satir As Long
    satir = Application.Caller.Row
    
    Dim ortalama As Double
    Dim kredi As Double
    
    ortalama = Cells(satir, 4).Value  ' 4. sütun = ort
    kredi = Cells(satir, 5).Value  ' 5. sütun = Kredi

    NotOrtKredi = ortalama * kredi
End Function

