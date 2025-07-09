Function NotOrt2() As Double
    Dim satir As Long
    satir = Application.Caller.Row
    
    Dim vize As Double
    Dim final As Double
    
    vize = Cells(satir, 2).Value
    final = Cells(satir, 3).Value
    NotOrt2 = vize * 0.3 + final * 0.7
End Function

