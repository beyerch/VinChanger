VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReadVin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Digital Horsepower, Inc VIN Editor : Read Vin"
   ClientHeight    =   3045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10320
      Top             =   2160
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2670
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar barCurrentTask 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar barReadVinTask 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblReadVinTask 
      Caption         =   "Read Vin Task :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label lblCurrentTask 
      Caption         =   "Current Task :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   9735
   End
End
Attribute VB_Name = "frmReadVin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
  'Clear Buffer
  varInBuffer = ""
  
End Sub

Private Sub Form_GotFocus()
   Dim intReturn As Integer
   
'Init things
  frmReadVin.Timer1.Interval = 1000
  frmReadVin.Timer1.Enabled = True
  
  'Init Time bars
  frmReadVin.barCurrentTask.Min = 0
  frmReadVin.barCurrentTask.Max = Module1.intMaxGenericTimer
  frmReadVin.Refresh
  
  'Read Vin
  'intScanToolType  As Integer '1 = atap 1.0, 2=atap 2.0, 3=dhp, 4=tech ii
  'MsgBox "scan tool = " & intScanToolType
  Select Case intScanToolType
     Case 1
        intReturn = ATv1_ReadVin
     Case 2
        intReturn = ReadVIN
     Case 3
        intReturn = DHPInterface_ReadVin
     Case Else
        MsgBox "You have not selected an interface cable yet.  Please select your cable and retry this operation."
        Unload Me
        frmMain.Visible = True
        
  End Select
     
  If intReturn = 0 Then
     'Store vin read
     strVinNumber = frmMain.Text1.Text
     'Disable box
     frmMain.cmdUpdateVIN.Enabled = False
  Else
     'Clear vin read
     strVinNumber = ""
     frmMain.Text1.Text = ""
     'Error occurred reading vin
     MsgBox "An error occurred while reading the VIN.  Please verify proper connection of interface and PCM!"
  End If
  
  frmReadVin.Visible = False
  Unload frmReadVin
  
End Sub

Private Sub Timer1_Timer()
  'Define variable for step of counter
  Dim intStatusIncrement As Double
  intStatusIncrement = 0.5
  'Update Counter Status Bar
  If frmReadVin.barCurrentTask.Value + intStatusIncrement <= frmReadVin.barCurrentTask.Max Then
     frmReadVin.barCurrentTask.Value = frmReadVin.barCurrentTask.Value + intStatusIncrement
  Else
     frmReadVin.barCurrentTask.Value = frmReadVin.barCurrentTask.Max
  End If
End Sub
