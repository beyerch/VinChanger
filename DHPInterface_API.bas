Attribute VB_Name = "DHPInterface_API"
Public Function DHPInterface_RequestFirmwareVersion()


  'Clear Buffer
  If ClearBuffer <> 0 Then
     DHPInterface_RequestFirmwareVersion = -1
     Exit Function
  End If
  
If boolAbortSendCommand = False Then
  'Reset Cable, by Setting Low Speed VPW Mode:
   rtnCode = DHPInterface_SetLowSpeedVPW
    Debug.Print "Set Low Speed VPW rtncode = " & rtnCode
  
    If Left(rtnCode, 1) = "-" Then
       'Return with error message
         DHPInterface_RequestFirmwareVersion = rtnCode
        Exit Function
    End If
End If

If boolAbortSendCommand = False Then
  'Set the cable to echo all instructdions that are sent to the bus.
    rtnCode = DHPInterface_SetMessageEcho
  If Left(rtnCode, 1) = "-" Then
     'Return with error message
     DHPInterface_RequestFirmwareVersion = rtnCode
     Exit Function
  End If
End If
  
If boolAbortSendCommand = False Then
  '--------------------Get Firmware  ------------------------
   'Display firmwave version

  'Send Get vin Part #1
  ReDim testArray(0)
  
  'Create Initialize packet
  testArray(0) = "&HB0"
  'Init Flags and Counters
  
  txtResponseString = DHPInterface_SendCommand("DHPInterface Request Firmware Version", testArray, 1, "&H92&H04")
  Debug.Print "DHPInterface Request Firmware Version = " & txtResponseString
   
        'Check response versus expected
        If Left(txtResponseString, 1) <> "-" Then
           DHPInterface_RequestFirmwareVersion = txtResponseString
         Else
            DHPInterface_RequestFirmwareVersion = txtResponseString
         End If

 
  
End If


'Close com port
If frmMain.MSComm1.PortOpen = True Then
    frmMain.MSComm1.PortOpen = False
End If


  
End Function


Public Function DHPInterface_WritePacket(arrayPacket() As Byte, arrayPacketLength)

'On Error GoTo WritePacketErrorHandler

'Check the data Packet and make sure it fits spec.

'Write data out
'frmMain.MSComm1.RTSEnable = True
'frmMain.MSComm1.DTREnable = True

frmMain.MSComm1.Output = arrayPacket()
   
DHPInterface_WritePacket = 0
Exit Function

WritePacketErrorHandler:
errValue = Err.Number
 errString = Err.Description
 errSource = Err.Source
 DHPInterface_WritePacket = Err.Number

End Function


Public Function DHPInterface_SetLowSpeedVPW()


 '----------------------------------------------------
  'Create Low Speed packet
  ReDim testArray(1)
  
  'Create Initialize packet
  testArray(0) = "&HC1"
  testArray(1) = "&H00"

  txtResponseString = DHPInterface_SendCommand("DHPInterface Set Low Speed VPW", testArray, 2, "&HC1&H00")
  Debug.Print "DHPInterface Set Low Speed VPW = " & txtResponseString
   
        'Check response versus expected
        If Left(txtResponseString, 1) <> "-" Then
           DHPInterface_SetLowSpeedVPW = "0"
        End If



End Function



Public Function DHPInterface_SetMessageEcho()


  '----------------------------------------------------
  'Enable Echo Mode
  ReDim testArray(2)
  
  'Create Initialize packet
  testArray(0) = "&H52"
  testArray(1) = "&H06"
  testArray(2) = "&H01"

  txtResponseString = DHPInterface_SendCommand("DHPInterface Set Message Echo", testArray, 3, "&H62&H06&H01")
  Debug.Print "DHPInterface Set Message Echo = " & txtResponseString
   
        'Check response versus expected
        If Left(txtResponseString, 1) <> "-" Then
           DHPInterface_SetMessageEcho = "0"
        End If



End Function

Public Function DHPInterface_ReadVin()


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
            DHPInterface_ReadVin = errValue
            Exit Function
        End If
  End If
  
  'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     DHPInterface_ReadVin = errValue
     Exit Function
  End If
  
  'Reset Cable, by Setting Low Speed VPW Mode:
  
   'Display Status Information for frmReadVin
   frmReadVin.lblCurrentTask.Caption = "Current Task : Resetting Cable"
   frmReadVin.barCurrentTask.Value = 0
   frmReadVin.barReadVinTask = 10
   frmReadVin.StatusBar1.Panels(1).Text = "Resetting Cable"
   frmReadVin.Refresh
  
    rtnCode = DHPInterface_SetLowSpeedVPW
    Debug.Print "Set Low Speed VPW rtncode = " & rtnCode
  If rtnCode <> 0 Then
     'Return with error message
     DHPInterface_ReadVin = rtnCode
     Exit Function
  End If
  
  'Set the cable to echo all instructdions that are sent to the bus.
  
   'Display Status Information for frmReadVin
   frmReadVin.lblCurrentTask.Caption = "Current Task : Configuring Communication Session"
   frmReadVin.barCurrentTask.Value = 0
   frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 10
   frmReadVin.StatusBar1.Panels(1).Text = "Configure Session"
   frmReadVin.Refresh
  
    rtnCode = DHPInterface_SetMessageEcho
  If rtnCode <> 0 Then
     'Return with error message
     DHPInterface_ReadVin = rtnCode
     Exit Function
  End If
  
   'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     DHPInterface_ReadVin = errValue
     Exit Function
  End If
  
  '--------------------Read VIN 1 ------------------------
  'Get the Vin Data starting with Part #1
   'Display Status Information for frmReadVin
   frmReadVin.lblCurrentTask.Caption = "Current Task : Reading VIN Data "
   frmReadVin.barCurrentTask.Value = 0
   frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 20
   frmReadVin.StatusBar1.Panels(1).Text = "Reading VIN Data"
   frmReadVin.Refresh
  
     
  'Send Get vin Part #1
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H05"
  testArray(1) = "&H6C"
  testArray(2) = "&H10"
  testArray(3) = "&HF1"
  testArray(4) = "&H3C"
  testArray(5) = "&H01"
  'Init Flags and Counters
  intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
  intReadRetries = 0
  boolDHPInterface_RequestFirmwareVersionPart1 = False
  Do Until boolDHPInterface_RequestFirmwareVersionPart1 = True Or intDHPInterface_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_ReadVin = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(testArray, 6)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
        response = DHPInterface_ReadPacket("&H0C&H00&H6C&HF1&H10&H7C&H01")
        If response <> "-1" Then
           Debug.Print "Vin 1 Read response = " & response
           boolDHPInterface_RequestFirmwareVersionPart1 = True
           intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
           txtvinpart1 = Chr(CDec(Mid(response, 33, 4))) & Chr(CDec(Mid(response, 37, 4))) & Chr(CDec(Mid(response, 41, 4))) & Chr(CDec(Mid(response, 45, 4))) & Chr(CDec(Mid(response, 49, 4)))
           frmMain.Refresh
           frmMain.Text1.Text = txtvinpart1
           
         Else
            intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
         End If
     Else
        intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_RequestFirmwareVersionPart1 = False Then
      DHPInterface_ReadVin = -2  'Failed Read part #1 of vin
      Exit Function
   End If
  '-------------END Read Vin #1---------------------------------------------------
    
    
    
  '--------------------Read VIN 2 ------------------------
  'Get the Vin Data starting with Part #2
   'Display Status Information for frmReadVin
   frmReadVin.lblCurrentTask.Caption = "Current Task : Reading VIN Data "
   frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 20
   frmReadVin.StatusBar1.Panels(1).Text = "Reading VIN Data"
   frmReadVin.Refresh
  
     
  'Send Get vin Part #2
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H05"
  testArray(1) = "&H6C"
  testArray(2) = "&H10"
  testArray(3) = "&HF1"
  testArray(4) = "&H3C"
  testArray(5) = "&H02"
  'Init Flags and Counters
  intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
  intReadRetries = 0
  boolDHPInterface_RequestFirmwareVersionPart1 = False
  Do Until boolDHPInterface_RequestFirmwareVersionPart1 = True Or intDHPInterface_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_ReadVin = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(testArray, 6)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
        response = DHPInterface_ReadPacket("&H0C&H00&H6C&HF1&H10&H7C&H02")
        If response <> "-1" Then
           Debug.Print "Vin 2 Read response = " & response
           boolDHPInterface_RequestFirmwareVersionPart1 = True
           intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
           txtvinpart1 = Chr(CDec(Mid(response, 29, 4))) & Chr(CDec(Mid(response, 33, 4))) & Chr(CDec(Mid(response, 37, 4))) & Chr(CDec(Mid(response, 41, 4))) & Chr(CDec(Mid(response, 45, 4))) & Chr(CDec(Mid(response, 49, 4)))
           frmMain.Refresh
           frmMain.Text1.Text = frmMain.Text1.Text & txtvinpart1
           
         Else
            intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
         End If
     Else
        intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_RequestFirmwareVersionPart1 = False Then
      DHPInterface_ReadVin = -2  'Failed Read part #1 of vin
      Exit Function
   End If
  '-------------END Read Vin #2---------------------------------------------------
    
    
    
  
  '--------------------Read VIN 3 ------------------------
  'Get the Vin Data starting with Part #3
   'Display Status Information for frmReadVin
   frmReadVin.lblCurrentTask.Caption = "Current Task : Reading VIN Data "
   frmReadVin.barReadVinTask = frmReadVin.barReadVinTask + 20
   frmReadVin.StatusBar1.Panels(1).Text = "Reading VIN Data"
   frmReadVin.Refresh
  
     
  'Send Get vin Part #3
  ReDim testArray(5)
  
  'Create Initialize packet
  testArray(0) = "&H05"
  testArray(1) = "&H6C"
  testArray(2) = "&H10"
  testArray(3) = "&HF1"
  testArray(4) = "&H3C"
  testArray(5) = "&H03"
  'Init Flags and Counters
  intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
  intReadRetries = 0
  boolDHPInterface_RequestFirmwareVersionPart1 = False
  Do Until boolDHPInterface_RequestFirmwareVersionPart1 = True Or intDHPInterface_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_ReadVin = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(testArray, 6)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
        response = DHPInterface_ReadPacket("&H0C&H00&H6C&HF1&H10&H7C&H03")
        If response <> "-1" Then
           Debug.Print "Vin 3 Read response = " & response
           boolDHPInterface_RequestFirmwareVersionPart1 = True
           intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
           txtvinpart1 = Chr(CDec(Mid(response, 29, 4))) & Chr(CDec(Mid(response, 33, 4))) & Chr(CDec(Mid(response, 37, 4))) & Chr(CDec(Mid(response, 41, 4))) & Chr(CDec(Mid(response, 45, 4))) & Chr(CDec(Mid(response, 49, 4)))
           frmMain.Refresh
           frmMain.Text1.Text = frmMain.Text1.Text & txtvinpart1
           
         Else
            intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
         End If
     Else
        intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_RequestFirmwareVersionPart1 = False Then
      DHPInterface_ReadVin = -3  'Failed Read part #2 of vin
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
      
      DHPInterface_ReadVin = 0
      
   End If
  '-------------END Read Vin #3---------------------------------------------------
    
'Close com port
frmMain.MSComm1.PortOpen = False
    



End Function
Public Function DHPInterface_ReadPacket(strExpectedAnswer As String) As String
On Error GoTo DHPInterface_ReadPacketErrorHandler
    'Declare Variables
    Dim boolPacketComplete As Boolean
    Dim intTimeout As Integer
    Dim intRetries As Integer
    Dim intMaxRetries As Integer
    Dim strTempBuffer As String
    
    
    'Initialize Variables
    boolPacketComplete = False
    boolTimedOut = False
    intTimeout = 1000 '1000 msec = 1 second
    intRetries = 0
    intMaxRetries = 3
    strTempBuffer = ""
    
    'Activate timer
    frmMain.Timer1.Enabled = True
    frmMain.Timer1.Interval = intTimeout
    
MainPacketLoop:
    Do While boolPacketComplete = False And intRetries < intMaxRetries
      DoEvents
       If boolTimedOut = True Then
            boolTimedOut = False
            intRetries = intRetries + 1
       Else
            intStartPos = InStr(1, varInBuffer, strExpectedAnswer)
            If intStartPos > 0 Then
               'Result matched return success and clear that item from buffer
               boolPacketComplete = True
               'Get data
               SOFhex = Left(strExpectedAnswer, 4)
               strTempBuffer = SOFhex
               Select Case Mid(SOFhex, 3, 1)
                 Case "0"
                   charFrameType = "0"
                    intDataLength = Val(SOFhex)
                 Case "2"
                    charFrameType = "2"
                    intDataLength = Val("&H" & Right(SOFhex, 1))
                 Case "3"
                    charFrameType = "3"
                    intDataLength = Val("&H" & Right(SOFhex, 1))
                 Case "6"
                    charFrameType = "6"
                    intDataLength = Val("&H" & Right(SOFhex, 1))
                 Case "8"
                    charFrameType = "8"
                    intDataLength = Val("&H" & Right(SOFhex, 1))
                 Case "9"
                    charFrameType = "9"
                    intDataLength = Val("&H" & Right(SOFhex, 1))
                 Case "C"
                    charFrameType = "C"
                    intDataLength = Val("&H" & Right(SOFhex, 1))
                 Case Else
                    SOFhex = ""
                    intDataLength = 0
                End Select
               
                For x = intStartPos + 4 To intStartPos + (intDataLength * 4) Step 4
                    strTempBuffer = strTempBuffer & Mid(varInBuffer, x, 4)
                Next x
                               
                'Buffer Debug.
                'Debug.Print "Attemping to Remove value from Buffer"
                'Debug.Print "Initial Buffer = " & varInBuffer
                'Debug.Print "Removing : " & strTempBuffer

                'Remove temp variable from buffer
                If intStartPos = 1 Then
                    If Len(varInBuffer) = Len(strTempBuffer) Then
                      varInBuffer = ""
                    Else
                      varInBuffer = Right(varInBuffer, Len(varInBuffer) - Len(strTempBuffer))
                    End If
                Else
                    varInBuffer = Left(varInBuffer, intStartPos - 1) & Mid(varInBuffer, intStartPos + Len(strTempBuffer), Len(varInBuffer) - Len(strTempBuffer) - 1)
                End If
                
                'Debug.Print "Buffer after removal = " & varInBuffer
                'Debug.Print "---------------------"
           End If
       End If
    Loop
        
    'De-Activate timer
    frmMain.Timer1.Enabled = False
    
    'Determine why we left the loop and act accordingly
    If boolPacketComplete = True Then
        DHPInterface_ReadPacket = strTempBuffer
    Else
        DHPInterface_ReadPacket = "-1"
    End If
    
    Exit Function

DHPInterface_ReadPacketErrorHandler:
    GoTo MainPacketLoop

End Function





Public Function DHPInterface_WriteVin()


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
            DHPInterface_WriteVin = errValue
            Exit Function
        End If
  End If
  
  'Clear Buffer
  If ClearBuffer <> 0 Then
     MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
     'Abort VIN Read
     DHPInterface_WriteVin = errValue
     Exit Function
  End If
  
  'Reset Cable, by Setting Low Speed VPW Mode:
  
   'Display Status Information for frmWriteVin
   frmWriteVin.lblCurrentTask.Caption = "Current Task : Resetting Cable"
   frmWriteVin.barCurrentTask.Value = 0
   frmWriteVin.barWriteVinTask = 10
   frmWriteVin.StatusBar1.Panels(1).Text = "Resetting Cable"
   frmWriteVin.Refresh
  
    rtnCode = DHPInterface_SetLowSpeedVPW
    Debug.Print "Set Low Speed VPW rtncode = " & rtnCode
  If rtnCode <> 0 Then
     'Return with error message
     DHPInterface_WriteVin = rtnCode
     Exit Function
  End If
  
  'Set the cable to echo all instructdions that are sent to the bus.
  
   'Display Status Information for frmWriteVin
   frmWriteVin.lblCurrentTask.Caption = "Current Task : Configuring Communication Session"
   frmWriteVin.barCurrentTask.Value = 0
   frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 10
   frmWriteVin.StatusBar1.Panels(1).Text = "Configure Session"
   frmWriteVin.Refresh
  
    rtnCode = DHPInterface_SetMessageEcho
  If rtnCode <> 0 Then
     'Return with error message
     DHPInterface_WriteVin = rtnCode
     Exit Function
  End If
    
    
  '--------------------Write VIN 1 ------------------------
  'Get the Vin Data starting with Part #1
   'Display Status Information for frmWriteVin
   frmWriteVin.lblCurrentTask.Caption = "Current Task : Writing VIN Data "
   frmWriteVin.barCurrentTask.Value = 0
   frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 20
   frmWriteVin.StatusBar1.Panels(1).Text = "Write VIN Data"
   frmWriteVin.Refresh
  
     
  'Send Get vin Part #1
  ReDim testArray(11)
  
  'Create Initialize packet
  testArray(0) = "&H0B"
  testArray(1) = "&H6C"
  testArray(2) = "&H10"
  testArray(3) = "&HF1"
  testArray(4) = "&H3B"
  testArray(5) = "&H01"
  testArray(6) = "&H00"
  testArray(7) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 1, 1))))
  testArray(8) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 2, 1))))
  testArray(9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 3, 1))))
  testArray(10) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 4, 1))))
  testArray(11) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 5, 1))))
  
  
  'Init Flags and Counters
  intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
  intReadRetries = 0
  boolDHPInterface_RequestFirmwareVersionPart1 = False
Do Until boolDHPInterface_RequestFirmwareVersionPart1 = True Or intDHPInterface_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_WriteVin = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(testArray, 12)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
     
        response = DHPInterface_ReadPacket("&H06&H00&H6C&HF1&H10&H7B&H01")
        If response <> "-1" Then
           Debug.Print "Write Vin 1 response = " & response
           boolDHPInterface_RequestFirmwareVersionPart1 = True
           intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
         Else
            intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
         End If
     Else
        intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_RequestFirmwareVersionPart1 = False Then
      DHPInterface_WriteVin = -2  'Failed Read part #1 of vin
      Exit Function
   End If
  '-------------END Write Vin #1---------------------------------------------------
    
    
    
    
    
  '--------------------Write VIN 2 ------------------------
  'Get the Vin Data starting with Part #2
   'Display Status Information for frmWriteVin
   frmWriteVin.lblCurrentTask.Caption = "Current Task : Writing VIN Data "
   frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 40
   frmWriteVin.StatusBar1.Panels(1).Text = "Write VIN Data"
   frmWriteVin.Refresh
  
     
  'Send Get vin Part #1
  ReDim testArray(11)
  
  'Create Initialize packet
  testArray(0) = "&H0B"
  testArray(1) = "&H6C"
  testArray(2) = "&H10"
  testArray(3) = "&HF1"
  testArray(4) = "&H3B"
  testArray(5) = "&H02"
  testArray(6) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 6, 1))))
  testArray(7) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 7, 1))))
  testArray(8) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 8, 1))))
  testArray(9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 9, 1))))
  testArray(10) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 10, 1))))
  testArray(11) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 11, 1))))
  
  
  'Init Flags and Counters
  intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
  intReadRetries = 0
  boolDHPInterface_RequestFirmwareVersionPart1 = False
Do Until boolDHPInterface_RequestFirmwareVersionPart1 = True Or intDHPInterface_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_WriteVin = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(testArray, 12)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
     
        response = DHPInterface_ReadPacket("&H06&H00&H6C&HF1&H10&H7B&H02")
        If response <> "-1" Then
           Debug.Print "Write Vin 1 response = " & response
           boolDHPInterface_RequestFirmwareVersionPart1 = True
           intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
         Else
            intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
         End If
     Else
        intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_RequestFirmwareVersionPart1 = False Then
      DHPInterface_WriteVin = -2  'Failed Read part #1 of vin
      Exit Function
   End If
  '-------------END Write Vin #2---------------------------------------------------
    
    
    
  '--------------------Write VIN 3 ------------------------
  'Write the vin data with section #3
   'Display Status Information for frmWriteVin
   frmWriteVin.lblCurrentTask.Caption = "Current Task : Writing VIN Data "
   frmWriteVin.barWriteVinTask = frmWriteVin.barWriteVinTask + 20
   frmWriteVin.StatusBar1.Panels(1).Text = "Write VIN Data"
   frmWriteVin.Refresh
  
     
  'Send Get vin Part #1
  ReDim testArray(11)
  
  'Create Initialize packet
  testArray(0) = "&H0B"
  testArray(1) = "&H6C"
  testArray(2) = "&H10"
  testArray(3) = "&HF1"
  testArray(4) = "&H3B"
  testArray(5) = "&H03"
  testArray(6) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 12, 1))))
  testArray(7) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 13, 1))))
  testArray(8) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 14, 1))))
  testArray(9) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 15, 1))))
  testArray(10) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 16, 1))))
  testArray(11) = "&H" & Hex(Asc(UCase(Mid(frmMain.Text1.Text, 17, 1))))
  
  
  'Init Flags and Counters
  intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
  intReadRetries = 0
  boolDHPInterface_RequestFirmwareVersionPart1 = False
Do Until boolDHPInterface_RequestFirmwareVersionPart1 = True Or intDHPInterface_RequestFirmwareVersionPart1Attempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_WriteVin = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(testArray, 12)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
     
        response = DHPInterface_ReadPacket("&H06&H00&H6C&HF1&H10&H7B&H03")
        If response <> "-1" Then
           Debug.Print "Write Vin 1 response = " & response
           boolDHPInterface_RequestFirmwareVersionPart1 = True
           intDHPInterface_RequestFirmwareVersionPart1Attempts = 0
         Else
            intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
         End If
     Else
        intDHPInterface_RequestFirmwareVersionPart1Attempts = intDHPInterface_RequestFirmwareVersionPart1Attempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_RequestFirmwareVersionPart1 = False Then
      DHPInterface_WriteVin = -3  'Failed Read part #1 of vin
      Exit Function
   End If
  '-------------END Write Vin #3---------------------------------------------------
    
    
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
   DHPInterface_WriteVin = 0
   
'close port
If frmMain.MSComm1.PortOpen = True Then
    frmMain.MSComm1.PortOpen = False
End If

   
End Function

Function DHPInterface_SendCommand(strCommandName, arrayPacket() As Byte, arrayPacketLength, strMatchPattern As String)


'Open Serial Port
  If frmMain.MSComm1.PortOpen = False Then
      If OpenSerialPort <> 0 Then
           DHPInterface_SendCommand = errValue
         Exit Function
      End If
  End If
  
'Init Flags and Counters
  intDHPInterface_SendCommandAttempts = 0
  intReadRetries = 0
  boolDHPInterface_SendCommand = False
  boolAbortSendCommand = False
  
  Do Until boolAbortSendCommand = True Or boolDHPInterface_SendCommand = True Or intDHPInterface_SendCommandAttempts > intMaxGenericRetries
    DoEvents
    'ClearBuffer
    If ClearBuffer <> 0 Then
       MsgBox "Error " & errValue & " has occurred while attempting to clear com buffer! " & vbCrLf & "Error Details : " & errString, vbCritical, "VinEditor : Com Buffer Initialization Error!"
       'Abort VIN Read
       DHPInterface_SendCommand = errValue
       Exit Function
    End If
    'Write test packet
     writepacketresponse = DHPInterface_WritePacket(arrayPacket, arrayPacketLength)
     'MsgBox "write packet response = " & writepacketresponse
     
     If writepacketresponse = 0 Then
     
        'If not expecting a response, bypass and return success
        If strMatchPattern = "" Then
            DHPInterface_SendCommand = 0
            Exit Function
        End If
            
     
        'Stop timer
        frmMain.Timer1.Enabled = False
     
        response = DHPInterface_ReadPacket(strMatchPattern)
        If response <> "-1" Then
           Debug.Print strCommandName & " response = " & response
           boolDHPInterface_SendCommand = True
           intDHPInterface_SendCommand = 0
           DHPInterface_SendCommand = response
         Else
            intDHPInterface_SendCommandAttempts = intDHPInterface_SendCommandAttempts + 1
         End If
     Else
        intDHPInterface_SendCommandAttempts = intDHPInterface_SendCommandAttempts + 1
     End If
   Loop
   'Check to see if failure occurred
   If boolDHPInterface_SendCommand = False Then
      DHPInterface_SendCommand = -2  'Failed Read of Send Command
      Exit Function
   End If

End Function


