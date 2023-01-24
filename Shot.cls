VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mRings As Double
Private mTeiler As Double
Private mXCoord As Integer
Private mYCoord As Integer

Public Sub Initialize(rings As Double, Teiler As Double, xCoord As Integer, yCoord As Integer)

    mRings = rings
    mTeiler = Teiler
    mXCoord = xCoord
    mYCoord = yCoord

End Sub

Public Property Get rings() As Variant
    rings = mRings
End Property

Public Property Get Teiler() As Variant
    Teiler = mTeiler
End Property

Public Property Get xCoord() As Variant
    xCoord = mXCoord
End Property

Public Property Get yCoord() As Variant
    yCoord = mYCoord
End Property
