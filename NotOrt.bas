Function NotOrt() As Double
    Dim satir As Long
    satir = Application.Caller.Row
    
    Dim vize As Double
    Dim final As Double
    
    vize = Cells(satir, 2).Value
    final = Cells(satir, 3).Value
    NotOrt = vize * 0.4 + final * 0.6
End Function
