Attribute VB_Name = "SAM4000Mock"
Public Function GetSerie(Scheiben As Integer) As Serie

    Const offset As Integer = 5

    Dim resultArr() As String
    
    'Generate String array with data
    'Array has some space which is not read and starts at an offset
    ReDim resultArr(offset + Scheiben * 4)
    resultArr(offset) = Scheiben
    For i = 0 To (Scheiben - 1)
        resultArr(offset + 1 + i * 4 + 0) = RandomRings()
        resultArr(offset + 1 + i * 4 + 1) = RandomTeiler()
        resultArr(offset + 1 + i * 4 + 2) = RandomCoordinates()
        resultArr(offset + 1 + i * 4 + 3) = RandomCoordinates()
    Next i
    
    MsgBox ("Mock einlegen, dann ok dr�cken")
    
    Set GetSerie = New Serie
    Call GetSerie.Initialize(resultArr)

End Function

Public Sub InitSAM()
    MsgBox ("Mock initialisiert")
End Sub

Public Function CreateBadSerie() As Serie
    Set CreateBadSerie = New Serie
    CreateBadSerie.Bad = True
End Function

'Values between -2000 and 2000
Private Function RandomCoordinates() As Integer
    RandomCoordinates = Int((2000 + 2000 + 1) * Rnd - 2000)
End Function

'Values between 0 and 10,9
Private Function RandomRings() As Single
    RandomRings = Round(10.9 * Rnd, 1)
End Function

'Values between 0 and 2800
Private Function RandomTeiler() As Single
    RandomTeiler = Round(2800 * Rnd, 1)
End Function
