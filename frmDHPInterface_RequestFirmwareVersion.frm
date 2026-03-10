VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDHPInterface_RequestFirmwareVersion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VinEditor 1.0 Beta : Get DHP Interface Properties"
   ClientHeight    =   4485
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requesting Interface Data ....."
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   5175
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame frmATap2Data 
      Caption         =   "DHP Interface Properties"
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5775
      Begin VB.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label1 
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Current Protocol :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Firmware Version :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Cable Type :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmDHPInterface_RequestFirmwareVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
   
 Dim rtncode
 
   frmDHPInterface_RequestFirmwareVersion.Refresh
   
  'MsgBox "get dhp cable info"

'Show status bar stuff
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
Me.Frame1.Visible = True
Me.ProgressBar1.Visible = True
Me.Timer1.Interval = 50
Me.Timer1.Enabled = True

Me.Refresh




rtncode = DHPInterface_RequestFirmwareVersion

Me.Timer1.Enabled = False
Me.Frame1.Visible = False
Me.ProgressBar1.Visible = False

If Left(rtncode, 1) <> "-" Then
  'Success
  'Parse into array
  
  'Update Status bar
  strCableConnected = "Connected "
  UpdateStatusBar
  frmMain.Refresh
  
  'Display Standard DataGet the Comm protocol
      frmDHPInterface_RequestFirmwareVersion.Label1(5) = "J1850 VPW"
      frmDHPInterface_RequestFirmwareVersion.Label1(1) = "DHP Flash Programming Cable 1.0"
  
   'Get Firmware version
  frmDHPInterface_RequestFirmwareVersion.Label1(3) = ResponsePacket(2)
  
Else
  'error
  MsgBox "Get properties failed"
End If

End Sub

Private Sub Form_Load()
Dim rtncode As Variant
Dim varYear

   frmDHPInterface_RequestFirmwareVersion.Left = ((Screen.Width - 100) / 2) - (frmDHPInterface_RequestFirmwareVersion.Width / 2)
   frmDHPInterface_RequestFirmwareVersion.Top = ((Screen.Height - 100) / 2) - (frmDHPInterface_RequestFirmwareVersion.Height / 2)


Me.Refresh




End Sub

Private Sub OKButton_Click()
Unload Me


End Sub

Private Sub Timer1_Timer()
If Me.ProgressBar1.Value + 10 <= Me.ProgressBar1.Max Then
  Me.ProgressBar1.Value = Me.ProgressBar1.Value + 10
Else
  Me.ProgressBar1.Value = Me.ProgressBar1.Min
End If
Me.Refresh

  
End Sub
