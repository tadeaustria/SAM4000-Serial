Attribute VB_Name = "SAM4000"
#If VBA7 Then
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const intPortID As Integer = 3 ' Ex. 1, 2, 3, 4 for COM1 - COM4

Public Function GetSerie(Scheiben As Integer) As Serie

    Dim data As String
        
    MsgBox ("Scheibe einlegen, dann ok drücken")
    data = Receive()
    Set GetSerie = CreateSerieWithSAMString(data)
    If GetSerie.Bad Then
        MsgBox ("Keine Daten gefunden")
    End If
    Do While GetSerie.NoOfShots < Scheiben And (Not GetSerie.Bad)
        Dim data2 As String
        Dim mySerie2 As Serie
        answer = MsgBox("Gefundene Scheiben " & GetSerie.NoOfShots & vbCrLf & " Zwischensumme: " & GetSerie.Sum & vbCrLf & " nächste Scheibe?", vbOKCancel)
        Select Case answer
        Case vbOK
            data2 = Receive()
            Set mySerie2 = CreateSerieWithSAMString(data2)
            If Not mySerie2.Bad Then
                Call GetSerie.Combine(mySerie2)
            End If
        Case vbCancel
            GetSerie.Bad = True
        End Select
    Loop

End Function

Public Sub InitSAM()

    Dim lngStatus As Long
    Dim strError  As String
    Dim strData   As String
    
    ' Initialize Communications
    lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID), _
        "baud=9600 parity=N data=8 stop=1")
    
    If lngStatus <> 0 Then
    ' Handle error.
        lngStatus = CommGetError(strError)
        MsgBox "COM Error: " & strError
        GoTo Error
    End If

    ' Set modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, True)
    lngStatus = CommSetLine(intPortID, LINE_DTR, True)
    
    Dim strDatat As String
    strData = HexToString("B205")

    ' Write data to serial port.
    lngSize = Len(strData)
    lngStatus = CommWrite(intPortID, strData)
    If lngStatus <> lngSize Then
        GoTo Error
    End If

    Dim result As String
    Dim result2 As String
    
    Sleep 500
    ' Read maximum of 64 bytes from serial port.
    lngStatus = CommRead(intPortID, result2, 1024)
    If lngStatus > 0 Then
        ' Process data.
        result = StringToHex(result2)
        If result = " 15" Then
            MsgBox ("Schnittstelle initialisiert")
            'Dim resArr() As String
            'resArr = Split(strData, Chr(13)) 'Split at CRLF
            'Dim mySerie As Serie
            'Set mySerie = CreateSerieWithSAMString(strData)
            'strData = HexToString("06") 'Ack
            'lngSize = Len(strData)
            'lngStatus = CommWrite(intPortID, strData)
            GoTo CloseCOM
        Else
            GoTo Error
        End If
    ElseIf lngStatus < 0 Then
        ' Handle error.
        GoTo Error
    End If

Error:
    MsgBox ("Fehler beim Initialisieren")

CloseCOM:
    ' Reset modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, False)
    lngStatus = CommSetLine(intPortID, LINE_DTR, False)

    ' Close communications.
    Call CommClose(intPortID)
End Sub

Public Function CreateBadSerie() As Serie
    Set CreateBadSerie = New Serie
    CreateBadSerie.Bad = True
End Function

Private Function Receive() As String

    Dim lngStatus As Long
    Dim strError  As String
    Dim strData   As String

    ' Initialize Communications
    lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID), _
        "baud=9600 parity=N data=8 stop=1")
    
    If lngStatus <> 0 Then
    ' Handle error.
        lngStatus = CommGetError(strError)
        MsgBox "COM Error: " & strError
        GoTo HandleError
    End If

    ' Set modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, True)
    lngStatus = CommSetLine(intPortID, LINE_DTR, True)
    
    Dim strDatat
    strData = HexToString("05")

    ' Write data to serial port.
    lngSize = Len(strData)
    lngStatus = CommWrite(intPortID, strData)
    If lngStatus <> lngSize Then
        GoTo HandleError
    End If

    Dim result As String
    Sleep 500
    ' Read maximum of 64 bytes from serial port.
    lngStatus = CommRead(intPortID, Receive, 1024)
    If lngStatus > 0 Then
        ' Process data.
        result = StringToHex(Receive)
        If result <> " 15" Then
            'Dim resArr() As String
            'resArr = Split(strData, Chr(13)) 'Split at CRLF
            'Dim mySerie As Serie
            'Set mySerie = CreateSerieWithSAMString(strData)
            strData = HexToString("06") 'Ack
            lngSize = Len(strData)
            lngStatus = CommWrite(intPortID, strData)
        Else
            GoTo HandleError
        End If
    ElseIf lngStatus < 0 Then
        ' Handle error.
        GoTo HandleError
    End If

    GoTo CloseCOM

HandleError:
    Receive = "-1"
CloseCOM:
    ' Reset modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, False)
    lngStatus = CommSetLine(intPortID, LINE_DTR, False)

    ' Close communications.
    Call CommClose(intPortID)
End Function

Private Function CreateSerieWithSAMString(SAMString As String) As Serie

    Dim splittedString() As String
    splittedString = Split(Replace(SAMString, ".", ","), Chr(13))
    Dim asd As Integer
    asd = UBound(splittedString, 1)
    If asd <> -1 Then
        If splittedString(0) <> "-1" Then
            Set CreateSerieWithSAMString = CreateSerie(splittedString)
            Exit Function
        End If
    End If
    
Error:
    Set CreateSerieWithSAMString = New Serie
    CreateSerieWithSAMString.Bad = True
End Function

Private Function CreateSerie(resultArr() As String) As Serie
    
    Set CreateSerie = New Serie
    Call CreateSerie.Initialize(resultArr)
    
End Function

Private Function HexToString(ByVal HexToStr As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim i         As Long
    For i = 1 To Len(HexToStr) Step 3
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, i, 2)))
        strReturn = strReturn & strTemp
    Next i
    HexToString = strReturn
End Function

Private Function StringToHex(ByVal StrToHex As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim i         As Long
    For i = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & Space$(1) & strTemp
    Next i
    StringToHex = strReturn
End Function
