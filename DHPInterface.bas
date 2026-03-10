Attribute VB_Name = "ATv1_API"
Public Function ATv1_RequestFirmwareVersion()


'INit variables
  intReadRetries = 0

  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      frmATv1_RequestFirmwareVersion.Refresh
      If OpenSerialPort <> 0 Then
         ATv1_RequestFirmwareVersion = -1
         Exit Function
      End If
  End If

  'Clear Buffer
  If ClearBuffer <> 0 Then
     ATv1_RequestFirmwareVersion = -1
     Exit Function
  End If
  
    '----------------------------------------------------
  'Send ATAP First call.
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H0b"
  testArray(2) = "&H05"
  testArray(3) = "&H73"
  testArray(4) = "&H00"
  testArray(5) = "&H00"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H00"
  testArray(12) = "&H80"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  Do Until boolATv1_RequestFirmwareVersionPart1 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = ATv1_WritePacket(testArray, 15)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 15
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ATv1_ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "read response = " & ReadPacketResponse
        
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H1" And _
           Right(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H4" Then
           'MsgBox " Success REsponse = " & MakeString(ResponsePacket(), ResponsePacketLength)
           boolATv1_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolATv1_RequestFirmwareVersionPart1 = True Or intReadRetries > intMaxGenericRetries
     Else
        intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv1_RequestFirmwareVersionPart1 = False Then
      ATv1_RequestFirmwareVersion = -1
      Exit Function
   End If
  '----------------------------------------------------------------
  
  
  
  
  'Clear Buffer
  If ClearBuffer <> 0 Then
     ATv1_RequestFirmwareVersion = -1
     Exit Function
  End If
  
    '----------------------------------------------------
  'Send ATAP Get parameters
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H68"
  testArray(2) = "&H6a"
  testArray(3) = "&Hf1"
  testArray(4) = "&H01"
  testArray(5) = "&H00"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H05"
  testArray(12) = "&H80"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  Do Until boolATv1_RequestFirmwareVersionPart1 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = ATv1_WritePacket(testArray, 15)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 15
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ATv1_ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "read response = " & ReadPacketResponse
        
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H1" And _
           Right(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H4" Then
           'MsgBox " Success REsponse = " & MakeString(ResponsePacket(), ResponsePacketLength)
           boolATv1_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolATv1_RequestFirmwareVersionPart1 = True Or intReadRetries > intMaxGenericRetries
     Else
        intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv1_RequestFirmwareVersionPart1 = False Then
      ATv1_RequestFirmwareVersion = -2
   Else
      ATv1_RequestFirmwareVersion = MakeString(ResponsePacket(), ResponsePacketLength)
   End If
  '----------------------------------------------------------------
  
  


  
End Function


Public Function ATv1_WritePacket(arrayPacket() As Byte, arrayPacketLength)

'On Error GoTo WritePacketErrorHandler

'Check the data Packet and make sure it fits spec.

'Check SOF
If arrayPacket(0) <> "&H01" Then
  'Bad Packet SOF
  ATv1_WritePacket = -1
  Exit Function
End If

'Make sure length of packet is correct
If arrayPacketLength <> 15 Then
   'Bad Packet Length
   ATv1_WritePacket = -2
   Exit Function
End If


'Write data out
frmMain.MSComm1.RTSEnable = True
frmMain.MSComm1.DTREnable = True

frmMain.MSComm1.Output = arrayPacket()
   
ATv1_WritePacket = 0
Exit Function

WritePacketErrorHandler:
errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 ATv1_WritePacket = Err.Number

End Function




Public Function ATv1_ReadVin()


  'INit variables
  intReadRetries = 0
  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
        'Display Status Information for frmReadVin
        frmReadVin.lblCurrentTask.Caption = "Current Task : Initializing Serial Port"
        frmReadVin.barCurrentTask.Value = 0
        frmReadVin.barReadVinTask = 10
        frmReadVin.StatusBar1.Panels(1).Text = "Initializing Serial Port"
        frmReadVin.Refresh
  
        'End Display Information
        If OpenSerialPort <> 0 Then
            MsgBox "Error " & errValue & " has occurred while attempting to open the com port! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Port Initialization Error!"
            'Abort VIN Read
            ATv1_ReadVin = errValue
            Exit Function
        End If
  End If
  
  'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     ATv1_ReadVin = errValue
     Exit Function
  End If
  
        'Display Status Information for frmReadVin
        frmReadVin.lblCurrentTask.Caption = "Current Task : Resetting Cable"
        frmReadVin.barCurrentTask.Value = 0
        frmReadVin.barReadVinTask = 10
        frmReadVin.StatusBar1.Panels(1).Text = "Resetting Cable"
        frmReadVin.Refresh
  
  '----------------------------------------------------
  'Send ATAP 1.0 Reset
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H00"
  testArray(2) = "&H00"
  testArray(3) = "&H00"
  testArray(4) = "&H00"
  testArray(5) = "&H00"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H00"
  testArray(12) = "&H40"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   'frmMain.Timer1.Interval = 25000
   frmMain.Timer1.Interval = 5000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   'Clear buffer
   ClearBuffer
   
  '--------------------------------------------------------
   
   
   
  'Display Status Information for frmReadVin
  frmReadVin.lblCurrentTask.Caption = "Current Task : Read Vin Data"
  frmReadVin.barCurrentTask.Value = 0
  frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 10
  frmReadVin.StatusBar1.Panels(1).Text = "Initializing Read Vin Data"
  frmReadVin.Refresh
  'End Display Infor/mation

  '----------------------------------------------------
    
    
  '----------------------------------------------------
  'Send Get vin Part #1
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H6c"
  testArray(2) = "&H10"
  testArray(3) = "&Hf1"
  testArray(4) = "&H3c"
  testArray(5) = "&H01"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H05"
  testArray(12) = "&H0f"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  Do Until boolATv1_RequestFirmwareVersionPart1 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = ATv1_WritePacket(testArray, 15)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 15
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ATv1_ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "read response = " & ReadPacketResponse
        
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H1" And _
           Right(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H4" Then
           'MsgBox " Success REsponse = " & MakeString(ResponsePacket(), ResponsePacketLength)
           txtvinpart1 = Chr$(ResponsePacket(7)) & Chr$(ResponsePacket(8)) & _
            Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(12))
           frmMain.Text1.Text = txtvinpart1
           frmMain.Refresh
           
           boolATv1_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolATv1_RequestFirmwareVersionPart1 = True Or intReadRetries > intMaxGenericRetries
     Else
        intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv1_RequestFirmwareVersionPart1 = False Then
      ATv1_ReadVin = -2  'Failed Read part #1 of vin
      Exit Function
   End If
  '----------------------------------------------------------------
    
   
'----------------------------------------------------
  'Display Status Information for frmReadVin
  frmReadVin.lblCurrentTask.Caption = "Current Task : Read Vin Data"
  frmReadVin.barCurrentTask.Value = 0
  frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 20
  frmReadVin.StatusBar1.Panels(1).Text = "Read Vin Data"
  frmReadVin.Refresh
  'End Display Infor/mation
    
'Send Get vin Part #2
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H6c"
  testArray(2) = "&H10"
  testArray(3) = "&Hf1"
  testArray(4) = "&H3c"
  testArray(5) = "&H02"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H05"
  testArray(12) = "&H0f"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  Do Until boolATv1_RequestFirmwareVersionPart1 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = ATv1_WritePacket(testArray, 15)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 15
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ATv1_ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "read response = " & ReadPacketResponse
        
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H1" And _
           Right(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H4" Then
           'MsgBox " Success REsponse = " & MakeString(ResponsePacket(), ResponsePacketLength)
           txtVinPart2 = Chr$(ResponsePacket(6)) & Chr$(ResponsePacket(7)) & Chr$(ResponsePacket(8)) & _
            Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(12))
           frmMain.Text1.Text = frmMain.Text1.Text & txtVinPart2
           frmMain.Refresh
           
           boolATv1_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolATv1_RequestFirmwareVersionPart1 = True Or intReadRetries > intMaxGenericRetries
     Else
        intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv1_RequestFirmwareVersionPart1 = False Then
      ATv1_ReadVin = -3  'Failed Read part #2 of vin
      Exit Function
   End If
  '----------------------------------------------------------------
    
    
'----------------------------------------------------
  'Display Status Information for frmReadVin
  frmReadVin.lblCurrentTask.Caption = "Current Task : Read Vin Data"
  frmReadVin.barCurrentTask.Value = 0
  frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 20
  frmReadVin.StatusBar1.Panels(1).Text = "Read Vin Data"
  frmReadVin.Refresh
  'End Display Infor/mation
    
    
'Send Get vin Part #3
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H6c"
  testArray(2) = "&H10"
  testArray(3) = "&Hf1"
  testArray(4) = "&H3c"
  testArray(5) = "&H03"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H05"
  testArray(12) = "&H0f"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  Do Until boolATv1_RequestFirmwareVersionPart1 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = ATv1_WritePacket(testArray, 15)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 15
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ATv1_ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "read response = " & ReadPacketResponse
        
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H1" And _
           Right(MakeString(ResponsePacket(), ResponsePacketLength), 3) = "&H4" Then
           'MsgBox " Success REsponse = " & MakeString(ResponsePacket(), ResponsePacketLength)
           txtVinPart3 = Chr$(ResponsePacket(6)) & Chr$(ResponsePacket(7)) & Chr$(ResponsePacket(8)) & _
            Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(12))
           frmMain.Text1.Text = frmMain.Text1.Text & txtVinPart3
           frmMain.Refresh
           boolATv1_RequestFirmwareVersionPart1 = True
           intReadRetries = 0
         Else
            intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolATv1_RequestFirmwareVersionPart1 = True Or intReadRetries > intMaxGenericRetries
     Else
        intATv1_RequestFirmwareVersionPart1Attempts = intATv1_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolATv1_RequestFirmwareVersionPart1 = False Then
      ATv1_ReadVin = -3  'Failed Read part #2 of vin
      Exit Function
   Else
      frmMain.Timer1.Enabled = False
      
      frmReadVin.barReadVinTask.Value = frmReadVin.barReadVinTask.Max
      frmReadVin.Refresh
      
      'We're done reading the vin
      strCableConnected = "Connected "
      strPCMConnected = "Connected "
      strVinNumber = frmMain.Text1.Text
      UpdateStatusBar
      frmMain.Refresh
      
      ATv1_ReadVin = 0
      Exit Function
      
   End If
  '----------------------------------------------------------------
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
   'Check to see if failure occurred
   If boolReadVinPart3 = False Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      ReadVIN = -4
      Exit Function
   Else
      frmMain.Timer1.Enabled = False
      
      frmReadVin.barReadVinTask.Value = frmReadVin.barReadVinTask.Max
      frmReadVin.Refresh
      
      'We're done reading the vin
      strCableConnected = "Connected "
      strPCMConnected = "Connected "
      strVinNumber = frmMain.Text1.Text
      UpdateStatusBar
      frmMain.Refresh
      
      ReadVIN = 0
      Exit Function
      
   End If
  '-----------------------------------

frmMain.MSComm1.PortOpen = False


End Function


Public Function ATv1_ReadPacket(arrayPacket() As Byte, arrayPacketLength)

'On Error GoTo ATv1_ReadPacketErrorHandler


  
  'INitialize timer
  frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
  frmMain.Timer1.Enabled = True
  boolTimedOut = False

'Debug.Print "---Begin : ATv1_ReadPacket ----"
Do
 DoEvents
   'See if anything in buffer
   If frmMain.MSComm1.InBufferCount > 0 Then
     'MsgBox "buffer > 0 "
     'Find Start of Frame Byte Must be $01
     SOFhex = ""
     Do
       DoEvents
       If frmMain.MSComm1.InBufferCount > 0 Then
        SOF = frmMain.MSComm1.Input
        'MsgBox "SOF = " & SOF
        SOFhex = Hex(Asc(SOF))
        If Len(SOFhex) = 1 Then SOFhex = "0" & SOFhex
       End If
     Loop Until SOFhex = "01" Or boolTimedOut = True
     SOFhex = "&H" & SOFhex
     'Debug.Print "SOFhex = " & SOFhex
        
     'Verify no timeout
     If boolTimedOut = True Then
        'Return with error
        ATv1_ReadPacket = -5 'Timed Out
        Exit Function
     End If
     
     'if not timed out get Control Length
     ControlLength = ""
     
     'get next 14 bytes
     'Read Bytes
     ReDim arrayBytes(14)
     For x = 1 To 14
      arrayBytes(x) = ""
        Do
          DoEvents
           If frmMain.MSComm1.InBufferCount > 0 Then
               tempByte = frmMain.MSComm1.Input
               tempByteHex = Hex(Asc(tempByte))
               If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
               tempByteHex = "&H" & tempByteHex

               arrayBytes(x) = tempByteHex
               'Debug.Print "Control Byte " & x & " = " & arrayControlBytes(x) & "."
           End If
        Loop Until arrayBytes(x) <> "" Or boolTimedOut = True
             'Verify we're not timed out
       If boolTimedOut = True Then
          ATv1_ReadPacket = -5
          Exit Function
       End If

     Next x
     
    'Verify we're not timed out
       If boolTimedOut = True Then
          ATv1_ReadPacket = -5
          Exit Function
       End If
    
    'Build Data Packet
    arrayPacketLength = 15
    ReDim arrayPacket(arrayPacketLength) As Byte
    
    arrayPacket(0) = SOFhex
    For x = 1 To 14
       arrayPacket(x) = arrayBytes(x)
    Next x
    
    'Check Last byte to make sure its 04
    'MsgBox "Byte 15 = " & arrayPacket(14) & "."
    If arrayPacket(14) <> "&H04" Then
       'Set bad checksum
       ATv1_ReadPacket = -6 'Bad Checksum
       Exit Function
    Else
       boolPacketComplete = True
    End If
    
End If

Loop Until boolPacketComplete = True Or boolTimedOut = True

'check if bool timed out or not
If boolPacketComplete = True Then
  ATv1_ReadPacket = 0
Else
  ATv1_ReadPacket = -5 'timed out or bad packet
End If

  
Exit Function

ATv1_ReadPacketErrorHandler:
errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 ATv1_ReadPacket = Err.Number




End Function




Public Function ATv1_WriteVin()


  'INit variables
  intReadRetries = 0
  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
        'Display Status Information for frmReadVin
        frmWriteVin.lblCurrentTask.Caption = "Current Task : Initializing Serial Port"
        frmWriteVin.barCurrentTask.Value = 0
        frmWriteVin.barWriteVinTask = 10
        frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Serial Port"
        frmWriteVin.Refresh
  
        'End Display Information
        If OpenSerialPort <> 0 Then
            MsgBox "Error " & errValue & " has occurred while attempting to open the com port! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Port Initialization Error!"
            'Abort VIN Read
            ATv1_WriteVin = errValue
            Exit Function
        End If
  End If
  
  'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     Exit Function
     ATv1_WriteVin = errValue
  End If
  
        'Display Status Information for frmReadVin
        frmWriteVin.lblCurrentTask.Caption = "Current Task : Resetting Cable"
        frmWriteVin.barCurrentTask.Value = 0
        frmWriteVin.barWriteVinTask = 10
        frmWriteVin.StatusBar1.Panels(1).Text = "Resetting Cable"
        frmWriteVin.Refresh
  
  '----------------------------------------------------
  '----------------------------------------------------
  'Send ATAP 1.0 Reset
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H00"
  testArray(2) = "&H00"
  testArray(3) = "&H00"
  testArray(4) = "&H00"
  testArray(5) = "&H00"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H00"
  testArray(12) = "&H40"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   boolTimedOut = False
   frmMain.Timer1.Interval = 25000
   'frmMain.Timer1.Interval = 5000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   'Clear buffer
   ClearBuffer
  
    
  '--------------------------------------------------------
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Write Vin Data"
  frmWriteVin.barCurrentTask.Value = 0
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  frmWriteVin.StatusBar1.Panels(1).Text = "Writing Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
  '----------------------------------------------------
    
  'Send ATAP 1.0 Reset
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H6c"
  testArray(2) = "&H10"
  testArray(3) = "&Hf1"
  testArray(4) = "&H3b"
  testArray(5) = "&H01"
  testArray(6) = "&H00"
  testArray(7) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 1, 1))))
  testArray(8) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 2, 1))))
  testArray(9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 3, 1))))
  testArray(10) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 4, 1))))
  testArray(11) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 5, 1))))
  testArray(12) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 5, 1))))
  testArray(13) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 5, 1))))
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   boolTimedOut = False
   frmMain.Timer1.Interval = 25000
   'frmMain.Timer1.Interval = 3000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   'Clear buffer
   ClearBuffer
   
   
  '----------------------------------------------------
  'Send ATAP 1.0 Reset
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H00"
  testArray(2) = "&H00"
  testArray(3) = "&H00"
  testArray(4) = "&H00"
  testArray(5) = "&H00"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H00"
  testArray(12) = "&H40"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   boolTimedOut = False
   frmMain.Timer1.Interval = 25000
   'frmMain.Timer1.Interval = 10000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   
   '--------------------------------------------------------
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Write Vin Data"
  frmWriteVin.barCurrentTask.Value = 0
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  frmWriteVin.StatusBar1.Panels(1).Text = "Writing Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
  '----------------------------------------------------
    
  'Write Vin Part #2
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H6c"
  testArray(2) = "&H10"
  testArray(3) = "&Hf1"
  testArray(4) = "&H3b"
  testArray(5) = "&H02"
  testArray(6) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 6, 1))))
  testArray(7) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 7, 1))))
  testArray(8) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 8, 1))))
  testArray(9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 9, 1))))
  testArray(10) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 10, 1))))
  testArray(11) = "&H0b"
  testArray(12) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 11, 1))))
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   boolTimedOut = False
   frmMain.Timer1.Interval = 25000
   'frmMain.Timer1.Interval = 3000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   'Clear buffer
   ClearBuffer
   
   '----------------------------------------------------
  'Send ATAP 1.0 Reset
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H00"
  testArray(2) = "&H00"
  testArray(3) = "&H00"
  testArray(4) = "&H00"
  testArray(5) = "&H00"
  testArray(6) = "&H00"
  testArray(7) = "&H00"
  testArray(8) = "&H00"
  testArray(9) = "&H00"
  testArray(10) = "&H00"
  testArray(11) = "&H00"
  testArray(12) = "&H40"
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   boolTimedOut = False
   frmMain.Timer1.Interval = 25000
   'frmMain.Timer1.Interval = 10000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
  
   
   
'--------------------------------------------------------
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Write Vin Data"
  frmWriteVin.barCurrentTask.Value = 0
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  frmWriteVin.StatusBar1.Panels(1).Text = "Writing Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
  '----------------------------------------------------
    
  'Write Vin Part #3
  ReDim testArray(14)
  
  'Create Initialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H6c"
  testArray(2) = "&H10"
  testArray(3) = "&Hf1"
  testArray(4) = "&H3b"
  testArray(5) = "&H03"
  testArray(6) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 12, 1))))
  testArray(7) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 13, 1))))
  testArray(8) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 14, 1))))
  testArray(9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 15, 1))))
  testArray(10) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 16, 1))))
  testArray(11) = "&H0b"
  testArray(12) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 17, 1))))
  testArray(13) = "&H00"
  testArray(14) = "&H04"

  boolATv1_RequestFirmwareVersionPart1 = False
  'Write test packet
  writepacketresponse = ATv1_WritePacket(testArray, 15)
  '----------------------------------------------------------------

   'Wait 5 seconds
   boolTimedOut = False
   frmMain.Timer1.Interval = 25000
   'frmMain.Timer1.Interval = 3000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   'Clear buffer
   ClearBuffer
   
   
      'stop timer
   frmMain.Timer1.Enabled = False
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Write Vin Data"
  frmWriteVin.barCurrentTask = frmWriteVin.barCurrentTask.Max
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask.Max
  
  frmWriteVin.StatusBar1.Panels(1).Text = "Write Vin Data Completed"
  frmWriteVin.Refresh
  'End Display Information
   
      
      MsgBox "Vin change successful.  Please turn key off to save new Vin."
      'We're done reading the vin
      strCableConnected = "Connected "
      strPCMConnected = "Connected "
      strVinNumber = frmMain.Text1.Text
      UpdateStatusBar
      frmMain.Refresh

   
   'Return Success
   ATv1_WriteVin = 0
   
'close port
frmMain.MSComm1.PortOpen = False

   
End Function

