Attribute VB_Name = "ATv2_API"


Public Function ATv2_RequestFirmwareVersion()

  'INit variables
  intReadRetries = 0

  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      frmATv2_RequestFirmwareVersion.Refresh
      If OpenSerialPort <> 0 Then
         ATv2_RequestFirmwareVersion = -1
         Exit Function
      End If
  End If

  'Clear Buffer
  If ClearBuffer <> 0 Then
     ATv2_RequestFirmwareVersion = -1
     Exit Function
  End If
  
  
  '----------------------------------------------------
  'Send ATAP Request Firmware Version
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H00"
  testArray(3) = "&H00"
  testArray(4) = "&H02"
 
   
  'Attempt to Request Firmware Version
  intATv2_RequestFirmwareVersionPart1Attempts = 0
  boolATv2_RequestFirmwareVersionPart1 = False
Do Until boolATv2_RequestFirmwareVersionPart1 = True Or intATv2_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H4&H80" Then
           ATv2_RequestFirmwareVersion = Chr$(Asc(ResponsePacket(3))) & "." & Chr$(Asc(ResponsePacket(4))) & Chr$(Asc(ResponsePacket(5)))
           boolATv2_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intReadRetries = intReadRetries + 1
         End If
      Loop While ReadPacketResponse > -1 And boolATv2_RequestFirmwareVersionPart1 = False
      If ReadPacketResponse < 0 Then
            intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
      End If
     Else
        intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv2_RequestFirmwareVersionPart1 = False Then
      ATv2_RequestFirmwareVersion = -1
      Exit Function
   End If
  '----------------------------------------------------------------
  
  'Close comport
  If frmMain.MSComm1.PortOpen = True Then
    frmMain.MSComm1.PortOpen = False
  End If
  
  
End Function





Public Function ATv2_RequestFirmwareDate()

  'INit variables
  intReadRetries = 0

  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      If OpenSerialPort <> 0 Then
         ATv2_RequestFirmwareDate = -1
         Exit Function
      End If
  End If

  'Clear Buffer
  If ClearBuffer <> 0 Then
     ATv2_RequestFirmwareDate = -1
     Exit Function
  End If
  
  
  '----------------------------------------------------
  'Send ATAP Request Firmware Version
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H01"
  testArray(3) = "&H00"
  testArray(4) = "&H03"
 
   
  'Attempt to Request Firmware Version
  intATv2_RequestFirmwareVersionPart1Attempts = 0
  boolATv2_RequestFirmwareVersionPart1 = False
Do Until boolATv2_RequestFirmwareVersionPart1 = True Or intATv2_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H4&H81" Then
           ATv2_RequestFirmwareDate = Chr$(Asc(ResponsePacket(4))) & "/" & Chr$(Asc(ResponsePacket(3))) & "/" & Chr$(Asc(ResponsePacket(5)))
           boolATv2_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intReadRetries = intReadRetries + 1
         End If
      Loop While ReadPacketResponse > -1 And boolATv2_RequestFirmwareVersionPart1 = False
      If ReadPacketResponse < 0 Then
            intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
      End If
     Else
        intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv2_RequestFirmwareVersionPart1 = False Then
      ATv2_RequestFirmwareDate = -1
      Exit Function
   End If
  '----------------------------------------------------------------
End Function




Public Function ATv2_RequestLDVModelNumber()

  'INit variables
  intReadRetries = 0

  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      If OpenSerialPort <> 0 Then
         ATv2_RequestLDVModelNumber = -1
         Exit Function
      End If
  End If

  'Clear Buffer
  If ClearBuffer <> 0 Then
     ATv2_RequestLDVModelNumber = -1
     Exit Function
  End If
  
Debug.Print "Request ATv2 Model Number"
  '----------------------------------------------------
  'Send ATAP Request Firmware Version
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H02"
  testArray(3) = "&H00"
  testArray(4) = "&H04"
 
   
  'Attempt to Request Firmware Version
  boolATv2_RequestFirmwareVersionPart1 = False
  intATv2_RequestFirmwareVersionPart1Attempts = 0
Do Until boolATv2_RequestFirmwareVersionPart1 = True Or intATv2_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     Debug.Print "Write Packet : " & MakeString(testArray, 5) & " Return Value = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 1 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        Debug.Print "Request ATv2 Model Number Response Packet = " & MakeString(ResponsePacket(), ResponsePacketLength)
        
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H1" Then
           ModelString = ""
           For x = 0 To Val(Chr$(Asc(ResponsePacket(1)))) - 2
              ModelString = ModelString & Chr$(ResponsePacket(3 + x))
           Next x
                     
           ATv2_RequestLDVModelNumber = ModelString
           boolATv2_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intReadRetries = intReadRetries + 1
         End If
      Loop While ReadPacketResponse > -1 And boolATv2_RequestFirmwareVersionPart1 = False
      If ReadPacketResponse < 0 Then
            intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
      End If
     Else
        intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv2_RequestFirmwareVersionPart1 = False Then
      ATv2_RequestLDVModelNumber = -1
      Exit Function
   End If
  '----------------------------------------------------------------
End Function




Public Function ATv2_RequestCurrentProtocol()

  'INit variables
  intReadRetries = 0

  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      If OpenSerialPort <> 0 Then
         ATv2_RequestCurrentProtocol = -1
         Exit Function
      End If
  End If

  'Clear Buffer
  If ClearBuffer <> 0 Then
     ATv2_RequestCurrentProtocol = -1
     Exit Function
  End If
  
  
  '----------------------------------------------------
  'Send ATAP Request Firmware Version
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H03"
  testArray(3) = "&H00"
  testArray(4) = "&H05"
 
   
  'Attempt to Request Firmware Version
  intATv2_RequestFirmwareVersionPart1Attempts = 0
  boolATv2_RequestFirmwareVersionPart1 = False
Do Until boolATv2_RequestFirmwareVersionPart1 = True Or intATv2_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 1 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H2&H83" Then
           Select Case Chr$(Asc(ResponsePacket(3)))
              Case "0"
                 ATv2_RequestCurrentProtocol = "No protocol selected"
              Case "1"
                 ATv2_RequestCurrentProtocol = "J1850 VPW"
              Case "2"
                 ATv2_RequestCurrentProtocol = "J1850 PWM"
              Case "3"
                 ATv2_RequestCurrentProtocol = "ISO9141-2"
            End Select
           boolATv2_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intReadRetries = intReadRetries + 1
         End If
      Loop While ReadPacketResponse > -1 And boolATv2_RequestFirmwareVersionPart1 = False
      If ReadPacketResponse < 0 Then
            intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
      End If
     Else
        intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv2_RequestFirmwareVersionPart1 = False Then
      ATv2_RequestCurrentProtocol = -1
      Exit Function
   End If
  '----------------------------------------------------------------

End Function



Public Function ATv2_RequestLDVSerial()

  'INit variables
  intReadRetries = 0

  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      If OpenSerialPort <> 0 Then
             'Abort Request Firmware Version
         ATv2_RequestLDVSerial = -1
         Exit Function
      End If
  End If

  'Clear Buffer
  If ClearBuffer <> 0 Then
     'Abort Request Firmware Version
     ATv2_RequestLDVSerial = -1
     Exit Function
  End If
  
  
  '----------------------------------------------------
  'Send ATAP Request Firmware Version
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H04"
  testArray(3) = "&H00"
  testArray(4) = "&H06"
 
   
  'Attempt to Request Firmware Version
  intATv2_RequestFirmwareVersionPart1Attempts = 0
  boolATv2_RequestFirmwareVersionPart1 = False
Do Until boolATv2_RequestFirmwareVersionPart1 = True Or intATv2_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 1 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&HB&H84" Then
           SerialNum = ""
           For x = 0 To 9
              SerialNum = SerialNum & Chr$(ResponsePacket(3 + x))
           Next x
                     
           ATv2_RequestLDVSerial = SerialNum
           boolATv2_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intReadRetries = intReadRetries + 1
         End If
      Loop While ReadPacketResponse > -1 And boolATv2_RequestFirmwareVersionPart1 = False
      If ReadPacketResponse < 0 Then
            intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
      End If
     Else
        intATv2_RequestFirmwareVersionPart1Attempts = intATv2_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv2_RequestFirmwareVersionPart1 = False Then
      ATv2_RequestLDVSerial = -1
      Exit Function
   End If
  '----------------------------------------------------------------
End Function

