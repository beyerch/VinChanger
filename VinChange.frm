VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Digital Horsepower, Inc VIN Utility                                                                                    "
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdateVIN 
      Caption         =   "&Update VIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdReadVIN 
      Caption         =   "&Read VIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2805
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   240
   End
   Begin VB.TextBox Text1 
      DataSource      =   "mscomm1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   8192
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Label Label1 
      Caption         =   "Vin Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDeviceConfig 
         Caption         =   "&Device Config"
      End
      Begin VB.Menu mnuCommConfig 
         Caption         =   "&Comm Config"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdReadVIN_Click()

    'Load and center help form
    frmReadVin.Left = ((Screen.Width - 100) / 2) - (frmReadVin.Width / 2)
    frmReadVin.Top = ((Screen.Height - 100) / 2) - (frmReadVin.Height / 2)
    frmReadVin.StatusBar1.Panels(1).Width = frmReadVin.StatusBar1.Width
    frmReadVin.Show vbModal, frmMain
     frmReadVin.Refresh
   
End Sub

Private Sub cmdUpdateVIN_Click()
'Load and center help form
   frmWriteVin.Left = ((Screen.Width - 100) / 2) - (frmWriteVin.Width / 2)
   frmWriteVin.Top = ((Screen.Height - 100) / 2) - (frmWriteVin.Height / 2)
   frmWriteVin.StatusBar1.Panels(1).Width = frmWriteVin.StatusBar1.Width
   
   frmWriteVin.Show vbModal, frmMain
   frmWriteVin.Refresh
   
  'MsgBox "VIN Update Complete!", vbOKOnly, "Digital Horsepower, Inc : Vin Update"
  
End Sub

Private Sub Form_Load()
'disable the update vin
  frmMain.cmdUpdateVIN.Enabled = False
  
  'Initialize
  boolComPortOpen = False
  strCableType = "Unspecified"
  strCableConnected = "Disconnected"
  strPCMConnected = "Disconnected"
  strVinNumber = "      -- None --      "
InitStatusBar
End Sub

Private Sub Form_Resize()
Dim x, Y   ' Declare variables.
'   If Form1.WindowState = vbMinimized Then
 '     Form1.Icon = LoadPicture("c:\myicon.ico")
  ' An icon named "myicon.ico" must be in the
  ' c:\ directory for this example to work
  ' correctly.
   'End If

End Sub

Private Sub Form_Terminate()
  'Close down comm port
  If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
  End If
  End
End Sub

Private Sub mnuAbout_Click()
   'Load and center help form
   frmAbout.Left = ((Screen.Width - 100) / 2) - (frmAbout.Width / 2)
   frmAbout.Top = ((Screen.Height - 100) / 2) - (frmAbout.Height / 2)
   frmAbout.Show vbModal, frmMain
   
End Sub

Private Sub mnuCommConfig_Click()
   'Load and center help form
   frmSerialDialog.Left = ((Screen.Width - 100) / 2) - (frmSerialDialog.Width / 2)
   frmSerialDialog.Top = ((Screen.Height - 100) / 2) - (frmSerialDialog.Height / 2)
   frmSerialDialog.Show vbModal, frmMain

End Sub

Private Sub mnuDeviceConfig_Click()
   'Load and center help form
   frmInterfaceDialog.Left = ((Screen.Width - 100) / 2) - (frmInterfaceDialog.Width / 2)
   frmInterfaceDialog.Top = ((Screen.Height - 100) / 2) - (frmInterfaceDialog.Height / 2)
   frmInterfaceDialog.Show vbModal, frmMain

End Sub

Private Sub mnuExit_Click()
  'Check and close com port
  If boolComPortOpen = True Then
     'Close com port
    frmMain.MSComm1.PortOpen = False
  End If
  
  End
  
  
  
End Sub

Private Sub mnuProperties_Click()
'intScanToolType  As Integer '1 = atap 1.0, 2=atap 2.0, 3=dhp, 4=tech ii
'MsgBox "scan tool = " & intScanToolType
Select Case intScanToolType
  Case 1
   'Load and center help form
   frmATv1_RequestFirmwareVersion.Left = ((Screen.Width - 100) / 2) - (frmATv1_RequestFirmwareVersion.Width / 2)
   frmATv1_RequestFirmwareVersion.Top = ((Screen.Height - 100) / 2) - (frmATv1_RequestFirmwareVersion.Height / 2)
   frmATv1_RequestFirmwareVersion.Show vbModal, frmMain
  Case 2
   'Load form
   frmATv2_RequestFirmwareVersion.Show vbModal, frmMain
  Case 3
   'Load  form
   frmDHPInterface_RequestFirmwareVersion.Show vbModal, frmMain
   
  Case Else
   MsgBox "You have not selected an interface cable yet.  Please select your cable and retry this operation."
   Unload Me
   frmMain.Visible = True
  
End Select


End Sub

Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of chars.
        Do While frmMain.MSComm1.InBufferCount > 0
                 tempByte = frmMain.MSComm1.Input
                 tempByteHex = Hex(Asc(tempByte))
                 If Len(tempByteHex) = 1 Then tempByteHex = "0" & tempByteHex
                 tempByteHex = "&H" & tempByteHex
                 varInBuffer = varInBuffer & tempByteHex
                 Debug.Print varInBuffer
        Loop

      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
   End Select

End Sub

Private Sub Text1_Change()
   'Compare the two
   frmMain.cmdUpdateVIN.Enabled = False
   
   If Len(Text1.Text) = 17 Then
      'Check to see if different than original
      If Trim(Text1.Text) <> Trim(strVinNumber) Then
         'No longer equal but different, allow write function
         frmMain.cmdUpdateVIN.Enabled = True
      End If
   End If
         
End Sub

Private Sub Timer1_Timer()
  'We've timed out set timed out
  boolTimedOut = True

End Sub
