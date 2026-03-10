Attribute VB_Name = "Module1"
Public errValue As Integer
Public errString As String
Public errSource As String

Public testArray() As Byte
Public ResponsePacket() As Byte
Public ResponsePacketLength As Integer





'Declare ScanTool Public Variables
Public intScanToolType  As Integer '1 = atap 1.0, 2=atap 2.0, 3=dhp, 4=tech ii
Public boolComPortOpenFlag As Boolean

Public strCableType As String
Public strCableConnected As String
Public strPCMConnected As String
Public strVinNumber As String
Public varInBuffer As Variant



Public boolTimedOut
'Declare Main value
Public Const intMaxGenericRetries = 3
Public Const intMaxJ1850Retries = 3
Public Const intMaxGenericTimer = 1
Public Const intMaxJ1850Timer = 1

Public Function ReadVIN()

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
     ReadVIN = errValue
     Exit Function
  End If
  End If
  'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     Exit Function
     ReadVIN = errValue
  End If
  
  'Send ATAP Reset Packet
  ReDim testArray(4)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H22"
  testArray(3) = "&H00"
  testArray(4) = "&H24"
  
  'Attempt to initialize
  boolInitAtap20 = False
  'Initialize STatus info
  frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 10
  strAttempts = ""
  intInitAtap20Attempts = 0
  
  Do
    'Check for retry
    If intInitAtap20Attempts > 0 Then
       strAttempts = " : Retry Attempt " & intInitAtap20Attempts
    End If
    'Display Status Information for frmReadVin
    frmReadVin.lblCurrentTask.Caption = "Current Task : Initializing Atap 2.0" & strAttempts
    frmReadVin.barCurrentTask.Value = 0
    frmReadVin.StatusBar1.Panels(1).Text = "Initializing Atap 2.0" & strAttempts
    frmReadVin.Refresh
    'End Display Information
  
    DoEvents
    ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     If writepacketresponse = 0 Then
        'Kill 2 seconds
        frmMain.Timer1.Interval = 3000
        frmMain.Timer1.Enabled = True
        Do
            DoEvents
        Loop Until boolTimedOut = True
     
        'Stop timer
        frmMain.Timer1.Enabled = False
        'Reset Interval
        frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer

     
        'Check for response.
        ReadPacketAttempts = 0
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
        Do
            ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
            'Debug.Print "Init Atap 2.0 Return Value = " & ReadPacketResponse
            'Check response versus expected
            If MakeString(ResponsePacket(), ResponsePacketLength) = "&H1&H1&HA2&H0&HA4" Then
                'Update Status Bar
                strCableConnected = "   Connected "
                boolInitAtap20 = True
                UpdateStatusBar
            Else
                strCableConnected = " Disconnected "
                'Standard packet problems keep cycling
                ReadPacketAttempts = ReadPacketAttempts + 1
            End If
        Loop Until boolInitAtap20 = True Or (ReadPacketAttempts > 1 And ReadPacketResponse = -5) Or ReadPacketAttempts > intMaxJ1850Retries
        
        'Check to see if too many read attemps
        If boolInitAtap20 = False Then
           'Error triggered logic exit increment counter
           intInitAtap20Attempts = intInitAtap20Attempts + 1
        End If
     Else
        intInitAtap20Attempts = intInitAtap20Attempts + 1
     End If
   Loop Until boolInitAtap20 = True Or intInitAtap20Attempts > intMaxGenericRetries
   'Check to see if failure occurred
   If boolInitAtap20 = False Then
      MsgBox "Unable to Initialize Autotap 2.0 Interface.  Aborting!"
      ReadVIN = -2 ' unable to interfacde 2.0 cable
      Exit Function
   End If
   

   'Wait 25 seconds
   'frmMain.Timer1.Interval = 25000
   frmMain.Timer1.Interval = 5
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
    'Send Get Vin Part 1 request
  ReDim testArray(9)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H5"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3C"
  testArray(8) = "&H1"
  testArray(9) = "&HF2"
  
   
  'Attempt to Read Part 1 Vin
  boolReadVinPart1 = False
  Do Until boolReadVinPart1 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 10)
     
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
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H1&HC1" Then
           'Update Status Bar
           txtvinpart1 = Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(11)) & _
            Chr$(ResponsePacket(12)) & Chr$(ResponsePacket(13)) & Chr$(ResponsePacket(14))
           frmMain.Text1.Text = txtvinpart1
           frmMain.Refresh
           'MsgBox txtvinpart1
           boolReadVinPart1 = True
           UpdateStatusBar
           intReadRetries = 0
         Else
            intReadVinPart1Attempts = intReadVinPart1Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolReadVinPart1 = True Or intReadRetries > intMaxGenericRetries
     Else
        intReadVinPart1Attempts = intReadVinPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolReadVinPart1 = False Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
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
    
    
    'Send Get Vin Part 2 request
  ReDim testArray(9)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H5"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3C"
  testArray(8) = "&H2"
  testArray(9) = "&HF3"
  
   
  'Attempt to Read Part 2 Vin
  boolReadVinPart2 = False
  Do Until boolReadVinPart2 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 10)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 2 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H1&HC1" Then
           'Update Status Bar
           txtVinPart2 = Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(11)) & _
            Chr$(ResponsePacket(12)) & Chr$(ResponsePacket(13)) & Chr$(ResponsePacket(14))
           frmMain.Text1.Text = frmMain.Text1.Text & txtVinPart2
           frmMain.Refresh
           'MsgBox txtVinPart2
           boolReadVinPart2 = True
           UpdateStatusBar
           intReadRetries = 0
         Else
            intReadVinPart2Attempts = intReadVinPart2Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolReadVinPart2 = True Or intReadRetries > intMaxGenericRetries
     Else
        intReadVinPart2Attempts = intReadVinPart2Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolReadVinPart2 = False Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      Exit Function
   End If
  '-----------------------------------
   
'----------------------------------------------------
  'Display Status Information for frmReadVin
  frmReadVin.lblCurrentTask.Caption = "Current Task : Read Vin Data"
  frmReadVin.barCurrentTask.Value = 0
  frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 20
  frmReadVin.StatusBar1.Panels(1).Text = "Read Vin Data"
  frmReadVin.Refresh
  'End Display Infor/mation
    
    
    'Send Get Vin Part 3request
  ReDim testArray(9)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H5"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3C"
  testArray(8) = "&H3"
  testArray(9) = "&HF4"
  
   
  'Attempt to Read Part 3Vin
  boolReadVinPart3 = False
  Do Until boolReadVinPart3 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 10)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 3 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H1&HC1" Then
           'Update Status Bar
           txtVinPart3 = Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(11)) & _
            Chr$(ResponsePacket(12)) & Chr$(ResponsePacket(13)) & Chr$(ResponsePacket(14))
           frmMain.Text1.Text = frmMain.Text1.Text & txtVinPart3
           frmMain.Refresh
           'MsgBox txtVinPart3
           boolReadVinPart3 = True
           UpdateStatusBar
           intReadRetries = 0
         Else
            intReadVinPart3Attempts = intReadVinPart3Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolReadVinPart3 = True Or intReadRetries > intMaxGenericRetries
     Else
        intReadVinPart3Attempts = intReadVinPart3Attempts + 1
     End If
   Loop
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
End Function


Public Function WriteVIN()

  'INit variables
  intReadRetries = 0
  'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
    'Display Status Information for frmWriteVin
    frmWriteVin.lblCurrentTask.Caption = "Current Task : Initializing Serial Port"
    frmWriteVin.barCurrentTask.Value = 0
    frmWriteVin.barWriteVinTask = 10
    frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Serial Port"
    frmWriteVin.Refresh
     'End Display Information
  
    If OpenSerialPort <> 0 Then
         MsgBox "Error " & errValue & " has occurred while attempting to open the com port! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Port Initialization Error!"
        'Abort VIN Read
        WriteVIN = Err.Value
        Exit Function
    End If
  End If
  'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     Exit Function
     WriteVIN = Err.Value
  End If
  
  'Send ATAP Reset Packet
  ReDim testArray(4)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H22"
  testArray(3) = "&H00"
  testArray(4) = "&H24"
  
  'Attempt to initialize
  boolInitAtap20 = False
  'Initialize STatus info
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  strAttempts = ""
  intInitAtap20Attempts = 0
  
  Do
    'Check for retry
    If intInitAtap20Attempts > 0 Then
       strAttempts = " : Retry Attempt " & intInitAtap20Attempts
    End If
    'Display Status Information for frmWriteVin
    frmWriteVin.lblCurrentTask.Caption = "Current Task : Initializing Atap 2.0" & strAttempts
    frmWriteVin.barCurrentTask.Value = 0
    frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Atap 2.0" & strAttempts
    frmWriteVin.Refresh
    'End Display Information
  
    DoEvents
    ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     If writepacketresponse = 0 Then
        'Kill 2 seconds
        frmMain.Timer1.Interval = 3000
        frmMain.Timer1.Enabled = True
        Do
            DoEvents
        Loop Until boolTimedOut = True
     
        'Stop timer
        frmMain.Timer1.Enabled = False
        'Reset Interval
        frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer

     
        'Check for response.
        ReadPacketAttempts = 0
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
        Do
            ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
            'Debug.Print "Init Atap 2.0 Return Value = " & ReadPacketResponse
            'Check response versus expected
            If MakeString(ResponsePacket(), ResponsePacketLength) = "&H1&H1&HA2&H0&HA4" Then
                'Update Status Bar
                strCableConnected = "   Connected "
                boolInitAtap20 = True
                UpdateStatusBar
            Else
                strCableConnected = " Disconnected "
                'Standard packet problems keep cycling
                ReadPacketAttempts = ReadPacketAttempts + 1
            End If
        Loop Until boolInitAtap20 = True Or (ReadPacketAttempts > 1 And ReadPacketResponse = -5) Or ReadPacketAttempts > intMaxJ1850Retries
        
        'Check to see if too many read attemps
        If boolInitAtap20 = False Then
           'Error triggered logic exit increment counter
           intInitAtap20Attempts = intInitAtap20Attempts + 1
        End If
     Else
        intInitAtap20Attempts = intInitAtap20Attempts + 1
     End If
   Loop Until boolInitAtap20 = True Or intInitAtap20Attempts > intMaxGenericRetries
   'Check to see if failure occurred
   If boolInitAtap20 = False Then
      MsgBox "Unable to Initialize Autotap 2.0 Interface.  Aborting!"
      WriteVIN = -2 ' unable to interfacde 2.0 cable
      Exit Function
   End If
   
   'Wait 15 seconds
   frmMain.Timer1.Interval = 15000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'Reset bool
   boolTimedOut = False
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
  frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Write  Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation

  '----------------------------------------------------
    'Send Get Vin Part 1 request
  ReDim testArray(15)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H0B"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3B"
  testArray(8) = "&H01"
  testArray(9) = "&H00"
  'get characters 1 - 6
  For x = 1 To 5
    testArray(x + 9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, x, 1))))
  Next x
  testArray(15) = ComputeChecksum(testArray, 16)
  'MsgBox "Vin part 1 tring = " & MakeString(testArray, 16)
   
 'Attempt to write Part 1 Vin
  writepacketresponse = -254
  Do Until writepacketresponse <> -254
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 16)
     'MsgBox "response = " & writepacketresponse
     
  Loop
   'Check to see if failure occurred
   If writepacketresponse <> 0 Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      WriteVIN = -2
      Exit Function
   End If
   
   
   'Wait 6 seconds
   frmMain.Timer1.Interval = 6000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
    
    
'---------------
'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     Exit Function
     WriteVIN = Err.Value
  End If
  
  'Send ATAP Reset Packet
  ReDim testArray(4)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H22"
  testArray(3) = "&H00"
  testArray(4) = "&H24"
  
  'Attempt to initialize
  boolInitAtap20 = False
  'Initialize STatus info
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  strAttempts = ""
  intInitAtap20Attempts = 0
  
  Do
    'Check for retry
    If intInitAtap20Attempts > 0 Then
       strAttempts = " : Retry Attempt " & intInitAtap20Attempts
    End If
    'Display Status Information for frmWriteVin
    frmWriteVin.lblCurrentTask.Caption = "Current Task : Initializing Atap 2.0" & strAttempts
    frmWriteVin.barCurrentTask.Value = 0
    frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Atap 2.0" & strAttempts
    frmWriteVin.Refresh
    'End Display Information
  
    DoEvents
    ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     If writepacketresponse = 0 Then
        'Kill 2 seconds
        frmMain.Timer1.Interval = 3000
        frmMain.Timer1.Enabled = True
        Do
            DoEvents
        Loop Until boolTimedOut = True
     
        'Stop timer
        frmMain.Timer1.Enabled = False
        'Reset Interval
        frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer

     
        'Check for response.
        ReadPacketAttempts = 0
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
        Do
            ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
            'Debug.Print "Init Atap 2.0 Return Value = " & ReadPacketResponse
            'Check response versus expected
            If MakeString(ResponsePacket(), ResponsePacketLength) = "&H1&H1&HA2&H0&HA4" Then
                'Update Status Bar
                strCableConnected = "   Connected "
                boolInitAtap20 = True
                UpdateStatusBar
            Else
                strCableConnected = " Disconnected "
                'Standard packet problems keep cycling
                ReadPacketAttempts = ReadPacketAttempts + 1
            End If
        Loop Until boolInitAtap20 = True Or (ReadPacketAttempts > 1 And ReadPacketResponse = -5) Or ReadPacketAttempts > intMaxJ1850Retries
        
        'Check to see if too many read attemps
        If boolInitAtap20 = False Then
           'Error triggered logic exit increment counter
           intInitAtap20Attempts = intInitAtap20Attempts + 1
        End If
     Else
        intInitAtap20Attempts = intInitAtap20Attempts + 1
     End If
   Loop Until boolInitAtap20 = True Or intInitAtap20Attempts > intMaxGenericRetries
   'Check to see if failure occurred
   If boolInitAtap20 = False Then
      MsgBox "Unable to Initialize Autotap 2.0 Interface.  Aborting!"
      WriteVIN = -2 ' unable to interfacde 2.0 cable
      Exit Function
   End If
   
   'Wait 15 seconds
   frmMain.Timer1.Interval = 15000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
      'Reset bool
   boolTimedOut = False

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
  frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Write  Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
  
  '----------------------------------------------------
    'Send Vin Part 2 request
  ReDim testArray(15)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H0B"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3B"
  testArray(8) = "&H02"
  'get characters 6 - 11
  For x = 6 To 11
    testArray(x + 3) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, x, 1))))
  Next x
  testArray(15) = ComputeChecksum(testArray, 16)
  'MsgBox "Vin part 2 tring = " & MakeString(testArray, 16)
   
 'Attempt to write Part 2 Vin
  writepacketresponse = -254
  Do Until writepacketresponse <> -254
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 16)
     'MsgBox "response = " & writepacketresponse
     
  Loop
   'Check to see if failure occurred
   If writepacketresponse <> 0 Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      WriteVIN = -2
      Exit Function
   End If
   
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Write Vin Data"
  frmWriteVin.barCurrentTask.Value = 0
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Write  Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
  
   'Wait 6 seconds
   frmMain.Timer1.Interval = 6000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
'----------------

  'Send ATAP Reset Packet
  ReDim testArray(4)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H22"
  testArray(3) = "&H00"
  testArray(4) = "&H24"
  
  'Attempt to initialize
  boolInitAtap20 = False
  'Initialize STatus info
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
  strAttempts = ""
  intInitAtap20Attempts = 0
  
  Do
    'Check for retry
    If intInitAtap20Attempts > 0 Then
       strAttempts = " : Retry Attempt " & intInitAtap20Attempts
    End If
    'Display Status Information for frmWriteVin
    frmWriteVin.lblCurrentTask.Caption = "Current Task : Initializing Atap 2.0" & strAttempts
    frmWriteVin.barCurrentTask.Value = 0
    frmWriteVin.StatusBar1.Panels(1).Text = "Initializing Atap 2.0" & strAttempts
    frmWriteVin.Refresh
    'End Display Information
  
    DoEvents
    ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 5)
     If writepacketresponse = 0 Then
        'Kill 2 seconds
        frmMain.Timer1.Interval = 3000
        frmMain.Timer1.Enabled = True
        Do
            DoEvents
        Loop Until boolTimedOut = True
     
        'Stop timer
        frmMain.Timer1.Enabled = False
        'Reset Interval
        frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer

     
        'Check for response.
        ReadPacketAttempts = 0
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
        Do
            ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
            'Debug.Print "Init Atap 2.0 Return Value = " & ReadPacketResponse
            'Check response versus expected
            If MakeString(ResponsePacket(), ResponsePacketLength) = "&H1&H1&HA2&H0&HA4" Then
                'Update Status Bar
                strCableConnected = "   Connected "
                boolInitAtap20 = True
                UpdateStatusBar
            Else
                strCableConnected = " Disconnected "
                'Standard packet problems keep cycling
                ReadPacketAttempts = ReadPacketAttempts + 1
            End If
        Loop Until boolInitAtap20 = True Or (ReadPacketAttempts > 1 And ReadPacketResponse = -5) Or ReadPacketAttempts > intMaxJ1850Retries
        
        'Check to see if too many read attemps
        If boolInitAtap20 = False Then
           'Error triggered logic exit increment counter
           intInitAtap20Attempts = intInitAtap20Attempts + 1
        End If
     Else
        intInitAtap20Attempts = intInitAtap20Attempts + 1
     End If
   Loop Until boolInitAtap20 = True Or intInitAtap20Attempts > intMaxGenericRetries
   'Check to see if failure occurred
   If boolInitAtap20 = False Then
      MsgBox "Unable to Initialize Autotap 2.0 Interface.  Aborting!"
      WriteVIN = -2 ' unable to interfacde 2.0 cable
      Exit Function
   End If
   
   'Wait 15 seconds
   frmMain.Timer1.Interval = 15000
   frmMain.Timer1.Enabled = True
   Do
      DoEvents
   Loop Until boolTimedOut = True
   'Reset bool
   boolTimedOut = False

   'stop timer
   frmMain.Timer1.Enabled = False
   'Reset Interval
    frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
   
   'Clear buffer
   ClearBuffer
   
  '--------------------------------------------------------

  
  
  
  
  
  
  
  '----------------------------------------------------
    'Send Get Vin Part 3 request
  ReDim testArray(15)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H0B"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3B"
  testArray(8) = "&H03"
  'get characters 12 - 17
  For x = 12 To 17
    testArray(x - 3) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, x, 1))))
  Next x
  testArray(15) = ComputeChecksum(testArray, 16)
  'MsgBox "Vin part 1 tring = " & MakeString(testArray, 16)
   
 'Attempt to write Part 1 Vin
  writepacketresponse = -254
  Do Until writepacketresponse <> -254
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 16)
     'MsgBox "response = " & writepacketresponse
     
  Loop
   'Check to see if failure occurred
   If writepacketresponse <> 0 Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      WriteVIN = -2
      Exit Function
   Else
   'stop timer
   frmMain.Timer1.Enabled = False
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Write Vin Data"
'  frmWriteVin.barCurrentTask.Value = 100
  frmWriteVin.barWriteVinTask = 100
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
   End If

End Function

Function Leftovers()
   
  '----------------------------------------------------------------
   
'----------------------------------------------------
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Read Vin Data"
  frmWriteVin.barCurrentTask.Value = 0
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 20
  frmWriteVin.StatusBar1.Panels(1).Text = "Read Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
    
    
    'Send Get Vin Part 2 request
  ReDim testArray(10)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H5"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3C"
  testArray(8) = "&H2"
  testArray(9) = "&HF3"
  
   
  'Attempt to Read Part 2 Vin
  boolwritevinPart2 = False
  Do Until boolwritevinPart2 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 10)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 2 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H1&HC1" Then
           'Update Status Bar
           txtVinPart2 = Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(11)) & _
            Chr$(ResponsePacket(12)) & Chr$(ResponsePacket(13)) & Chr$(ResponsePacket(14))
           frmMain.Text1.Text = frmMain.Text1.Text & txtVinPart2
           frmMain.Refresh
           'MsgBox txtVinPart2
           boolwritevinPart2 = True
           UpdateStatusBar
           intReadRetries = 0
         Else
            intwritevinPart2Attempts = intwritevinPart2Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolwritevinPart2 = True Or intReadRetries > intMaxGenericRetries
     Else
        intwritevinPart2Attempts = intwritevinPart2Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolwritevinPart2 = False Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      Exit Function
   End If
  '-----------------------------------
   
'----------------------------------------------------
  'Display Status Information for frmWriteVin
  frmWriteVin.lblCurrentTask.Caption = "Current Task : Read Vin Data"
  frmWriteVin.barCurrentTask.Value = 0
  frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 20
  frmWriteVin.StatusBar1.Panels(1).Text = "Read Vin Data"
  frmWriteVin.Refresh
  'End Display Infor/mation
    
    
    'Send Get Vin Part 3request
  ReDim testArray(10)
  
  'Create INitialize packet
  testArray(0) = "&H01"
  testArray(1) = "&H01"
  testArray(2) = "&H41"
  testArray(3) = "&H5"
  testArray(4) = "&H6C"
  testArray(5) = "&H10"
  testArray(6) = "&HF1"
  testArray(7) = "&H3C"
  testArray(8) = "&H3"
  testArray(9) = "&HF4"
  
   
  'Attempt to Read Part 3Vin
  boolwritevinPart3 = False
  Do Until boolwritevinPart3 = True
    DoEvents
    'ClearBuffer
    'Write test packet
     writepacketresponse = WritePacket(testArray, 10)
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
      Do
        'Check for response.
        ResponsePacketLength = 255
        ReDim ResponsePacket(ResponsePacketLength)
  
        ReadPacketResponse = ReadPacket(ResponsePacket(), ResponsePacketLength)
        'MsgBox "Read Packet Error = " & ReadPacketResponse
        'MsgBox "Read Vin 3 string = " & MakeString(ResponsePacket(), ResponsePacketLength)
        'Check response versus expected
        If Left(MakeString(ResponsePacket(), ResponsePacketLength), 10) = "&H1&H1&HC1" Then
           'Update Status Bar
           txtVinPart3 = Chr$(ResponsePacket(9)) & Chr$(ResponsePacket(10)) & Chr$(ResponsePacket(11)) & _
            Chr$(ResponsePacket(12)) & Chr$(ResponsePacket(13)) & Chr$(ResponsePacket(14))
           frmMain.Text1.Text = frmMain.Text1.Text & txtVinPart3
           frmMain.Refresh
           'MsgBox txtVinPart3
           boolwritevinPart3 = True
           UpdateStatusBar
           intReadRetries = 0
         Else
            intwritevinPart3Attempts = intwritevinPart3Attempts + 1
            intReadRetries = intReadRetries + 1
         End If
      Loop Until boolwritevinPart3 = True Or intReadRetries > intMaxGenericRetries
     Else
        intwritevinPart3Attempts = intwritevinPart3Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolwritevinPart3 = False Then
      MsgBox "Unable to Communicate with PCM. Aborting!"
      WriteVIN = -4
      Exit Function
   Else
      frmMain.Timer1.Enabled = False
      
      frmWriteVin.barWriteVinTask.Value = frmWriteVin.barWriteVinTask.Max
      frmWriteVin.Refresh
      
      'We're done reading the vin
      strCableConnected = "Connected "
      strPCMConnected = "Connected "
      strVinNumber = frmMain.Text1.Text
      UpdateStatusBar
      frmMain.Refresh
      
      WriteVIN = 0
      Exit Function
      
   End If
  '-----------------------------------
End Function
Function leftovercrap()
  
   
   ' 9600 baud, no parity, 8 data, and 1 stop bit.
   MSComm1.Settings = "19200,N,8,1"
   ' Tell the control to read entire buffer when Input
   ' is used.
   MSComm1.InputLen = 1
   ' Open the port.
   MSComm1.PortOpen = True
   ' Send the attention command to the modem.
   'MSComm1.Output = "$01$01$22$00$24" & Chr$(13)
   Dim DataOutput() As Byte
   ReDim DataOutput(5)
If (MSComm1.InBufferCount > 0) Then
  Do While MSComm1.InBufferCount > 0
    tempinput = MSComm1.Input
  Loop
End If
   
   
'--------------reboot LDV
   DataOutput(0) = &H1
   DataOutput(1) = &H1
   DataOutput(2) = &H22
   DataOutput(3) = &H0
   DataOutput(4) = &H24
   
      MSComm1.Output = DataOutput()
      
   
   ' the modem responds with "OK".
   ' Wait for data to come back to the serial port.
   Text1.Text = ""
   'set input mode
   'MSComm1.InputMode = comInputModeBinary
hexString = ""
BytesRead = 0
Do
   datain = MSComm1.Input
   If Len(datain) > 0 Then
     BytesRead = BytesRead + 1
     temphex = Hex(Asc(datain))
     If Len(temphex) = 1 Then temphex = "0" & temphex
     hexString = hexString & "$" & temphex
   End If
   DoEvents

Loop Until BytesRead = 5
'MsgBox "Reboot STatus= " & hexString
'---------end reboot


If (MSComm1.InBufferCount > 0) Then
  Do While MSComm1.InBufferCount > 0
    tempinput = MSComm1.Input
  Loop
End If

   DataOutput(0) = &H1
   DataOutput(1) = &H1
   DataOutput(2) = &H13
   DataOutput(3) = &H0
   DataOutput(4) = &H15
   
      MSComm1.Output = DataOutput()
      
   
   ' the modem responds with "OK".
   ' Wait for data to come back to the serial port.
   Text1.Text = ""
   'set input mode
   'MSComm1.InputMode = comInputModeBinary
hexString = ""
BytesRead = 0
Do
   datain = MSComm1.Input
   If Len(datain) > 0 Then
     BytesRead = BytesRead + 1
     temphex = Hex(Asc(datain))
     If Len(temphex) = 1 Then temphex = "0" & temphex
     hexString = hexString & "$" & temphex
   End If
   DoEvents

Loop Until BytesRead = 6
'MsgBox "Supported Modes = " & hexString
   ' Read the "OK" response data in the serial port.
   ' Close the serial port.
   
   
   
If (MSComm1.InBufferCount > 0) Then
  Do While MSComm1.InBufferCount > 0
    tempinput = MSComm1.Input
  Loop
End If
   
   
   DataOutput(0) = &H1
   DataOutput(1) = &H1
   DataOutput(2) = &H3
   DataOutput(3) = &H0
   DataOutput(4) = &H15
   
      MSComm1.Output = DataOutput()
hexString = ""
Do
   datain = MSComm1.Input
   If Len(datain) > 0 Then
     temphex = Hex(Asc(datain))
     If Len(temphex) = 1 Then temphex = "0" & temphex
     hexString = hexString & "$" & temphex
   End If
   DoEvents

Loop Until Len(hexString) >= 18
'MsgBox "Current Mode = " & hexString
   
   
If (MSComm1.InBufferCount > 0) Then
  Do While MSComm1.InBufferCount > 0
    tempinput = MSComm1.Input
  Loop
End If


'change protocol to j1850 vpw
 '  ReDim DataOutput(6)
 '  DataOutput(0) = &H1
 '  DataOutput(1) = &H2
 '  DataOutput(2) = &H23
 '  DataOutput(3) = &H5
 '  DataOutput(4) = &H0
  ' DataOutput(5) = &H2B
   
   '   MSComm1.Output = DataOutput()
   
'hexString = ""
'BytesRead = 0
'Do
 '  datain = MSComm1.Input
 '  If Len(datain) > 0 Then
 '    BytesRead = BytesRead + 1
 '    temphex = Hex(Asc(datain))
  '   If Len(temphex) = 1 Then temphex = "0" & temphex
  '   hexString = hexString & "$" & temphex
  ' End If
  ' DoEvents

'Loop Until BytesRead = 6
'MsgBox "Change Protocol = " & hexString
   
   
   
   
'If (MSComm1.InBufferCount > 0) Then
'  Do While MSComm1.InBufferCount > 0
'    tempinput = MSComm1.Input
'  Loop
'End If
   
'reinit bus
'   ReDim DataOutput(5)
'   DataOutput(0) = &H1
'   DataOutput(1) = &H1
'   DataOutput(2) = &H21
'   DataOutput(3) = &H0
'   DataOutput(4) = &H23
'
'
'      MSComm1.Output = DataOutput()
'
'hexString = ""
'BytesRead = 0
'Do
'   datain = MSComm1.Input
'   If Len(datain) > 0 Then
'     BytesRead = BytesRead + 1
'     temphex = Hex(Asc(datain))
'     If Len(temphex) = 1 Then temphex = "0" & temphex
'     hexString = hexString & "$" & temphex
'   End If
'   DoEvents''

'Loop Until BytesRead = 5
'MsgBox "REinit = " & hexString''


'If (MSComm1.InBufferCount > 0) Then
'  Do While MSComm1.InBufferCount > 0
'    tempinput = MSComm1.Input
'  Loop
'End If


'change to 10.4K
  ' ReDim DataOutput(5)
  ' DataOutput(0) = &H1
  ' DataOutput(1) = &H1
  ' DataOutput(2) = &H25
  ' DataOutput(3) = &H0
  ' DataOutput(4) = &H27
   
   
   '   MSComm1.Output = DataOutput()
   
'hexString = ""
'BytesRead = 0
'Do
'   datain = MSComm1.Input
'   If Len(datain) > 0 Then
'     BytesRead = BytesRead + 1
'     temphex = Hex(Asc(datain))
 '    If Len(temphex) = 1 Then temphex = "0" & temphex
 '    hexString = hexString & "$" & temphex
 '  End If
  ' DoEvents

'Loop Until BytesRead = 5
'MsgBox "Low Speed J1850 = " & hexString

MSComm1.OutBufferCount = 0



'read Block 1 of vin
   ReDim DataOutput(10)
   DataOutput(0) = &H1
   DataOutput(1) = &H1
   DataOutput(2) = &H41
   DataOutput(3) = &H5
   DataOutput(4) = &H6C
   DataOutput(5) = &H10
   DataOutput(6) = &HF1
   DataOutput(7) = &H3C
   DataOutput(8) = &H1
   DataOutput(9) = &HF2
     
   
      MSComm1.Output = DataOutput()
   
hexString = ""
BytesRead = 0
Do
   datain = MSComm1.Input
   If Len(datain) > 0 Then
     BytesRead = BytesRead + 1
     temphex = Hex(Asc(datain))
     If Len(temphex) = 1 Then temphex = "0" & temphex
     hexString = hexString & "$" & temphex
     'Debug.Print hexString
   End If
   DoEvents

Loop Until MSComm1.InBufferCount < 1 And BytesRead > 0


'MsgBox "block reading = " & hexString

   'Write Block 1 of vin
   ReDim DataOutput(16)
   DataOutput(0) = &H1
   DataOutput(1) = &H1
   DataOutput(2) = &H41
   DataOutput(3) = &HB
   DataOutput(4) = &H6C
   DataOutput(5) = &H10
   DataOutput(6) = &HF1
   DataOutput(7) = &H3B
   DataOutput(8) = &H1
   DataOutput(9) = &H0
   DataOutput(10) = &H53
   DataOutput(11) = &H47
   DataOutput(12) = &H32
   DataOutput(13) = &H57
   DataOutput(14) = &H50
   DataOutput(15) = &H6A
     
   
      MSComm1.Output = DataOutput()
   
hexString = ""
BytesRead = 0
Do
   datain = MSComm1.Input
   If Len(datain) > 0 Then
     BytesRead = BytesRead + 1
     temphex = Hex(Asc(datain))
     If Len(temphex) = 1 Then temphex = "0" & temphex
     hexString = hexString & "$" & temphex
     'Debug.Print hexString
   End If
   DoEvents

Loop Until MSComm1.InBufferCount < 1 And BytesRead > 0

'MsgBox "block writing = " & hexString
   


'read Block 1 of vin
   ReDim DataOutput(10)
   DataOutput(0) = &H1
   DataOutput(1) = &H1
   DataOutput(2) = &H41
   DataOutput(3) = &H5
   DataOutput(4) = &H6C
   DataOutput(5) = &H10
   DataOutput(6) = &HF1
   DataOutput(7) = &H3C
   DataOutput(8) = &H1
   DataOutput(9) = &HF2
     
   
      MSComm1.Output = DataOutput()
   
hexString = ""
BytesRead = 0
Do
   datain = MSComm1.Input
   If Len(datain) > 0 Then
     BytesRead = BytesRead + 1
     temphex = Hex(Asc(datain))
     If Len(temphex) = 1 Then temphex = "0" & temphex
     hexString = hexString & "$" & temphex
     'Debug.Print hexString
   End If
   DoEvents

Loop Until MSComm1.InBufferCount < 1 And BytesRead > 0
'MsgBox "block reading = " & hexString

End Function

Private Sub Form_Load()
  frmMain.Visible = True
  frmMain.Enabled = True
  
  'refresh form
  frmMain.Refresh
  
  'Activeate status bar
  InitStatusBar

  'disable the update vin
  frmMain.cmdUpdateVIN.Enabled = False
  
  'Initialize
  boolComPortOpen = False
  strCableType = "Autotap 2.0"
  strCableConnected = "Disconnected"
  strPCMConnected = "Disconnected"
  strVinNumber = " -- None -- "
  
  'Do main loop
  Do
     DoEvents
     
     UpdateStatusBar
  Loop
  
 
  frmMain.Refresh
  
   


  




End Sub
Public Function MakeString(arrayPacket() As Byte, arrayPacketLength)

'make a string
MakeString = ""
For x = 0 To arrayPacketLength - 1
  MakeString = MakeString & "&H" & Hex(arrayPacket(x))
Next x


End Function
Public Function WritePacket(arrayPacket() As Byte, arrayPacketLength)

'On Error GoTo WritePacketErrorHandler

'Check the data Packet and make sure it fits spec.

'Check SOF
If arrayPacket(0) <> "&H01" Then
  'Bad Packet SOF
  WritePacket = -1
  Exit Function
End If

'Make sure length of packet is correct
If arrayPacketLength <> 1 + 1 + Val(MakeHex(arrayPacket(1))) + 1 + Val(MakeHex(arrayPacket(2 + Val(MakeHex(arrayPacket(1)))))) + 1 Then
   'Bad Packet Length
   WritePacket = -2
   Exit Function
End If

'Verify checksum
calculatedChecksum = ComputeChecksum(arrayPacket(), arrayPacketLength)
If MakeHex(calculatedChecksum) <> "&H" & MakeHex(Hex(arrayPacket(arrayPacketLength - 1))) Then
  'Bad Checksum
  WritePacket = -3
  Exit Function
End If

'Write data out

frmMain.MSComm1.Output = arrayPacket()
   
WritePacket = 0
Exit Function

WritePacketErrorHandler:
errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 WritePacket = Err.Number

End Function
Public Function MakeHex(strHex)
'Check for &H
If InStr(1, strHex, "$") Then
  MakeHex = "&H" & Right(strHex, Len(strHex) - 1)
Else
  MakeHex = strHex
End If



End Function

Public Function ComputeChecksum(arrayPacket() As Byte, arrayPacketLength)

Dim tempChecksumCalc
tempChecksumCalc = 0
For x = 0 To arrayPacketLength - 2
  tempChecksumCalc = tempChecksumCalc + Val(MakeHex(arrayPacket(x)))
Next x

'check to see if bigger than 255....
If tempChecksumCalc > 255 Then
   'return only last two hex bytes
   ComputeChecksum = "&H" & Right(Hex(tempChecksumCalc), 2)
Else
   ComputeChecksum = "&H" & Hex(tempChecksumCalc)
End If




End Function

Public Function ReadPacket(arrayPacket() As Byte, arrayPacketLength)

'On Error GoTo ReadPacketErrorHandler

  'Declare variables
  ReDim arrayControlBytes(255)
  ReDim arrayDataBytes(255)

  
  
  'INitialize timer
  frmMain.Timer1.Interval = 1000 * intMaxJ1850Timer
  frmMain.Timer1.Enabled = True
  boolTimedOut = False

'Debug.Print "---Begin : ReadPacket ----"
Do
 DoEvents
   'See if anything in buffer
   If frmMain.MSComm1.InBufferCount > 0 Then
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
        'Reset timeout
        boolTimedOut = False
        'Return with error
        ReadPacket = -5 'Timed Out
        Exit Function
     End If
     
     'if not timed out get Control Length
     ControlLength = ""
     
     Do
      DoEvents
        If frmMain.MSComm1.InBufferCount > 0 Then
            tempByte = frmMain.MSComm1.Input
            tempByteHex = Hex(Asc(tempByte))
            If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
            tempByteHex = "&H" & tempByteHex
            ControlLength = tempByteHex
            'Debug.Print "Control Length = " & ControlLength
        End If
     Loop Until ControlLength <> "" Or boolTimedOut = True
     
     'Verify that control length is between 1 and 32
     If Val(ControlLength) < 1 Or Val(ControlLength) > 32 Then
       'Set Error and Return
        ReadPacket = -3  'Bad control length
        Exit Function
     End If
     
     'Verify no timeout
     If boolTimedOut = True Then
        'Reset timeout
        boolTimedOut = False
        'Return with error
        ReadPacket = -5 'Timed Out
        Exit Function
     End If
     'Read Control Bytes
     ReDim arrayControlBytes(Val(ControlLength))
     For x = 1 To Val(ControlLength)
      arrayControlBytes(x) = ""
        Do
          DoEvents
           If frmMain.MSComm1.InBufferCount > 0 Then
               tempByte = frmMain.MSComm1.Input
               tempByteHex = Hex(Asc(tempByte))
               If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
               tempByteHex = "&H" & tempByteHex

               arrayControlBytes(x) = tempByteHex
               'Debug.Print "Control Byte " & x & " = " & arrayControlBytes(x) & "."
           End If
        Loop Until arrayControlBytes(x) <> "" Or boolTimedOut = True
             'Verify we're not timed out
     'Verify no timeout
     If boolTimedOut = True Then
        'Reset timeout
        boolTimedOut = False
        'Return with error
        ReadPacket = -5 'Timed Out
        Exit Function
     End If

     Next x
     
     'Verify no timeout
     If boolTimedOut = True Then
        'Reset timeout
        boolTimedOut = False
        'Return with error
        ReadPacket = -5 'Timed Out
        Exit Function
     End If
     
     
     
     'if not timed out get Data Length
     DataLength = ""
     Do
       DoEvents
        If frmMain.MSComm1.InBufferCount > 0 Then
           tempByte = frmMain.MSComm1.Input
           tempByteHex = Hex(Asc(tempByte))
           If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
           tempByteHex = "&H" & tempByteHex

           DataLength = tempByteHex
           'Debug.Print "Data Length = " & DataLength
        End If
     Loop Until DataLength <> "" Or boolTimedOut = True
          'Verify we're not timed out
     'Verify no timeout
     If boolTimedOut = True Then
        'Reset timeout
        boolTimedOut = False
        'Return with error
        ReadPacket = -5 'Timed Out
        Exit Function
     End If

     
     'Verify that data length is between 1 and 32
     If Val(DataLength) < 0 Or Val(DataLength) > 32 Then
       'Set Error and Return
        ReadPacket = -4  'Bad data length
        Exit Function
     End If
     
     'Read data Bytes
   If Val(DataLength) > 0 Then
     ReDim arrayDataBytes(Val(DataLength))
     For x = 1 To Val(DataLength)
      arrayDataBytes(x) = ""
        Do
          DoEvents
           If frmMain.MSComm1.InBufferCount > 0 Then
               tempByte = frmMain.MSComm1.Input
               tempByteHex = Hex(Asc(tempByte))
               If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
               tempByteHex = "&H" & tempByteHex
           
               arrayDataBytes(x) = tempByteHex
               'Debug.Print "Data Byte " & x & " = " & arrayDataBytes(x) & "."
           End If
        Loop Until arrayDataBytes(x) <> "" Or boolTimedOut = True
     'Verify we're not timed out
     'Verify no timeout
     If boolTimedOut = True Then
        'Reset timeout
        boolTimedOut = False
        'Return with error
        ReadPacket = -5 'Timed Out
        Exit Function
     End If
     
     Next x
  Else
    'Debug.Print "No Data Bytes"
  End If
    'Read checksum
    Do
     DoEvents
      'See if anything in buffer
      If frmMain.MSComm1.InBufferCount > 0 Then
         'Get Checksum
            tempByte = frmMain.MSComm1.Input
            tempByteHex = Hex(Asc(tempByte))
            If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
            tempByteHex = "&H" & tempByteHex

         checksum = tempByteHex
         'Debug.Print "Checksum = " & checksum
      End If
    Loop Until checksum <> "" Or boolTimedOut = True
     'Verify we're not timed out
       If boolTimedOut = True Then
          ReadPacket = -5
          boolTimeOut = False
       End If
     
    
    'Build Data Packet
    arrayPacketLength = Val(DataLength) + Val(ControlLength) + 4
    ReDim arrayPacket(arrayPacketLength) As Byte
    
    arrayPacket(0) = SOFhex
    arrayPacket(1) = ControlLength
    For x = 1 To Val(ControlLength)
       arrayPacket(x + 1) = arrayControlBytes(x)
    Next x
    arrayPacket(x + 1) = DataLength
    Y = 0
   If Val(DataLength) > 0 Then
    For Y = 1 To Val(DataLength)
       arrayPacket(x + Y + 1) = arrayDataBytes(Y)
    Next Y
   End If
    arrayPacket(Val(DataLength) + Val(ControlLength) + 3) = checksum

    'Compare checksum's
    If ComputeChecksum(arrayPacket(), Val(DataLength) + Val(ControlLength) + 4) <> Val(checksum) Then
       'Set bad checksum
       ReadPacket = -6 'Bad Checksum
       Exit Function
    Else
       boolPacketComplete = True
    End If
    
End If

Loop Until boolPacketComplete = True Or boolTimedOut = True

'check if bool timed out or not
If boolPacketComplete = True Then
  ReadPacket = 0
Else
  ReadPacket = -5 'timed out or bad packet
End If

  
Exit Function

ReadPacketErrorHandler:
errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 ReadPacket = Err.Number




End Function


Public Function ClearBuffer()

On Error GoTo ClearBufferErrorHandler

varInBuffer = ""

If (frmMain.MSComm1.InBufferCount > 0) Then
  'INitialize timer
  frmMain.Timer1.Interval = 1000 * intMaxGenericTimer
  frmMain.Timer1.Enabled = True
  boolTimedOut = False
  
  'MsgBox frmMain.MSComm1.InBufferCount
  
  Do While frmMain.MSComm1.InBufferCount > 0 And boolTimedOut = False
   DoEvents
    tempinput = frmMain.MSComm1.Input
  Loop
  
  'Disable Timer
  frmMain.Timer1.Enabled = False
  
  'return status
  If frmMain.MSComm1.InBufferCount < 1 And boolTimedOut = False Then
     ClearBuffer = 0
  Else
     errValue = -1
     errString = "Serial Timeout Occured while clearing Buffer!"
     errSource = "Clear Buffer"
     ClearBuffer = -1
  End If
Else
  ClearBuffer = 0
End If

Exit Function

ClearBufferErrorHandler:
errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 ClearBuffer = Err.Number


End Function

Public Function OpenSerialPort()

On Error GoTo SerialPortErrorHandler
   'Set com opbject ot one byte at a time processing
   frmMain.MSComm1.InputLen = 1
   ' Open the port.
   frmMain.MSComm1.PortOpen = True
If intScanToolType = 3 Then
   'Set to 57600 for dhp cable
   frmMain.MSComm1.Settings = "57600,N,8,1"
Else
   'set to 19200 for atap's
   frmMain.MSComm1.Settings = "19200,N,8,1"
End If
   frmMain.MSComm1.NullDiscard = False
   'Return 0 status
   OpenSerialPort = 0
   Exit Function

SerialPortErrorHandler:
 'One or more errors occurred, get error and return
 errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 OpenSerialPort = Err.Number


End Function

Public Function UpdateStatusBar()
   'Update values
   frmMain.StatusBar1.Panels(1).Text = "Cable Type : " & strCableType
   frmMain.StatusBar1.Panels(2).Text = "Cable Status : " & strCableConnected
   frmMain.StatusBar1.Panels(3).Text = "PCM Status : " & strPCMConnected
   frmMain.StatusBar1.Panels(4).Text = "Vin : " & strVinNumber


End Function
Public Function InitStatusBar()
 'Capture Errors
' On Error GoTo StatusBarErrorHandler
   'Set number of objecst to 4 (cable type, cable status, pcm status, vin)
   'Remove all objects
   frmMain.StatusBar1.Panels.Clear
   
   'Add a panel for Cable Type
   frmMain.StatusBar1.Panels.Add
   frmMain.StatusBar1.Panels(1).Text = "Cable Type : " & strCableType
   frmMain.StatusBar1.Panels(1).Alignment = sbrLeft
frmMain.StatusBar1.Panels(1).MinWidth = 90 * Len(frmMain.StatusBar1.Panels(1).Text)
   frmMain.StatusBar1.Panels(1).ToolTipText = "Current Cable Detected by Program"
   
   'Add a panel for Cable Status
   frmMain.StatusBar1.Panels.Add 2, , "Cable Status : " & strCableConnected
   frmMain.StatusBar1.Panels(2).Alignment = sbrLeft
   frmMain.StatusBar1.Panels(2).MinWidth = 90 * Len(frmMain.StatusBar1.Panels(2).Text)
   frmMain.StatusBar1.Panels(2).ToolTipText = "This denotes the current connection status of your cable."
   
   'Add PCM Status
   frmMain.StatusBar1.Panels.Add 3, , "PCM Status : " & strPCMConnected
   frmMain.StatusBar1.Panels(3).Alignment = sbrLeft
   frmMain.StatusBar1.Panels(3).MinWidth = 90 * Len(frmMain.StatusBar1.Panels(3).Text)
   frmMain.StatusBar1.Panels(3).ToolTipText = "This Denotes whether you are currently connected to a pcm or not."
   
   'Add vin #
   frmMain.StatusBar1.Panels.Add 4, , "Vin : " & strVinNumber
   frmMain.StatusBar1.Panels(4).Alignment = sbrLeft
   frmMain.StatusBar1.Panels(4).MinWidth = 90 * Len(frmMain.StatusBar1.Panels(4).Text)
   frmMain.StatusBar1.Panels(4).ToolTipText = "Current VIN read from the PCM."
   
   'Even out the widths if there is still space left
   intPanelWidths = frmMain.StatusBar1.Panels(1).MinWidth + frmMain.StatusBar1.Panels(2).MinWidth + frmMain.StatusBar1.Panels(3).MinWidth + frmMain.StatusBar1.Panels(4).MinWidth
   intRemainingWidth = frmMain.StatusBar1.Width - intPanelWidths
   If intRemainingWidth > 0 Then
      'Have spaceleft over Calculate and divide it evenly
      dblQuarterWidth = intRemainingWidth / frmMain.StatusBar1.Panels.Count
      'Increment each min widht
      For Each StatusPanel In frmMain.StatusBar1.Panels
          StatusPanel.MinWidth = StatusPanel.MinWidth + dblQuarterWidth
      Next
   End If
   
   frmMain.StatusBar1.Visible = True
   
   'Return Success
   InitStatusBar = 0
   Exit Function
   
StatusBarErrorHandler:
 'One or more errors occurred, get error and return
 errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 InitStatusBar = Err.Number
 
   

End Function





Private Function InitSerial(intPort As Integer, boolNullDiscard As Boolean, strBaudRate As String, strParity As String, strDataBits As String, strStopBits As String, intInputLength As Integer)

'Capture Errors
On Error GoTo errSerialHandler

'Set active com port
MSComm1.CommPort = intPort
'Set nulldiscard
MSComm1.NullDiscard = boolNullDiscard
'Set baud,parity,data, stop bits
MSComm1.Settings = strBaudRate & "," & strParity & "," & strDataBits & "," & strStopBits
'Tell control how many bytes to return with each buffer inquiry.
MSComm1.InputLen = intInputLength
'Return success
InitSerial = 0
Exit Function

errSerialHandler:
'One or more errors occurred, get error and return
errValue = Err.Number
errString = Err.Description
errSource = Err.Source
InitSerial = Err.Number


End Function
