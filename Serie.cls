VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Serie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mShots() As Shot
Private mNoOfShots As Integer
Private mBad As Boolean

Public Sub Initialize(resultArr() As String)
    
    Const offset As Integer = 5
    
    mNoOfShots = CInt(resultArr(offset))
    ReDim mShots(mNoOfShots - 1)
    For i = 0 To (mNoOfShots - 1)
        Dim shotArr(3) As String
        For j = 0 To 3
            shotArr(j) = resultArr(offset + 1 + j + (i * 4))
        Next j
        Set mShots(i) = CreateShot(shotArr)
    Next i
    mBad = False
    
End Sub

Public Property Get Bad() As Boolean
    Bad = mBad
End Property

Public Property Let Bad(ByVal value As Boolean)
    mBad = value
End Property

Public Property Get Shot(ByVal index As Integer) As Shot
    Set Shot = mShots(index)
End Property

Public Property Get MinTeiler() As Double
    MinTeiler = 9999.9
    For Each lShot In mShots
        If MinTeiler > lShot.Teiler Then
            MinTeiler = lShot.Teiler
        End If
    Next
End Property

Public Property Get NoOfShots() As Integer
    NoOfShots = mNoOfShots
End Property

Public Property Get Sum() As Double
    Sum = 0
    For Each lShot In mShots
        Sum = Sum + lShot.rings
    Next
End Property

Private Function CreateShot(result() As String) As Shot

    Set CreateShot = New Shot
    Call CreateShot.Initialize(CDbl(result(0)), convTeiler(result(1)), convCoord(result(2), result(1)), convCoord(result(3), result(1)))

End Function

Private Function convTeiler(Text As String) As Double
    On Error GoTo Error
    convTeiler = CDbl(Text)
    Exit Function
Error:
    convTeiler = 9999.9
End Function

Private Function convCoord(Text As String, TeilerText As String) As Integer
    'Koordinaten sind auch bei Fehler -1, daher Teiler zum Test verwenden
    On Error GoTo Error
    Dim test As Double
    test = CDbl(TeilerText)
    convCoord = CInt(Text)
    Exit Function
Error:
    convCoord = 9999
End Function

Public Sub Combine(otherSerie As Serie)
    Dim oldNoOfShots As Integer
    oldNoOfShots = mNoOfShots
    mNoOfShots = mNoOfShots + otherSerie.NoOfShots
    ReDim Preserve mShots(mNoOfShots - 1)
    
    For i = 0 To otherSerie.NoOfShots - 1
        Set mShots(oldNoOfShots + i) = otherSerie.Shot(i)
    Next i
End Sub

