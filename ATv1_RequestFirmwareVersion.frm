VERSION 5.00
Begin VB.Form frmATv1_RequestFirmwareVersion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VinEditor 1.0 Beta : Get Autotap 1.0 Properties"
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
   Begin VB.Frame frmATap2Data 
      Caption         =   "Autotap 1.0 Properties"
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      Begin VB.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Current Protocol :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Firmware Date :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Firmware Version :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
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
Attribute VB_Name = "frmATv1_RequestFirmwareVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Dim rtnCode As Variant
Dim varYear
'MsgBox "get atap v1 cable info"

rtnCode = ATv1_RequestFirmwareVersion

If Left(rtnCode, 1) <> "-" Then
  'Success
  'Parse into array
  
  'Update Status bar
  strCableConnected = "Connected "
  UpdateStatusBar
  frmMain.Refresh
  
  'Get the Comm protocol
  Select Case ResponsePacket(1)
    Case 1
      Me.Label1(5) = "J1850 VPW"
      Me.Label1(1) = "Autotap AT1"
    Case 2
      Me.Label1(5) = "J1850 PWM"
      Me.Label1(1) = "Autotap AT2"
    Case 3
      Me.Label1(5) = "ISO 9141-2"
      Me.Label1(1) = "Autotap AT3"
    Case Else
      Me.Label1(5) = "Unknown"
  End Select
  
  'Get the Date
  Me.Label1(4) = ResponsePacket(3) & "/" & ResponsePacket(4) & "/"
  If ResponsePacket(2) > 5 Then
     Me.Label1(4) = Me.Label1(4) & "19" & ResponsePacket(2)
  Else
    Me.Label1(4) = Me.Label1(4) & "20" & ResponsePacket(2)
  End If
    
  
  'Get Firmware version
  Me.Label1(3) = ResponsePacket(5) & "." & ResponsePacket(6) & ResponsePacket(7)
  
  'Get CAble Model
  If ResponsePacket(8) <> 0 Then Me.Label1(1) = ResponsePacket(8)
Else
  'error
  MsgBox "Get properties failed"
End If


End Sub

Private Sub OKButton_Click()
Unload Me


End Sub
