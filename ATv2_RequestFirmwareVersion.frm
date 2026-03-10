VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmATv2_RequestFirmwareVersion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VinEditor 1.0 Beta : Autotap 2.0  Properties"
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
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requesting Interface Data ....."
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   5175
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame frmATap2Data 
      Caption         =   "Autotap 2.0 Properties"
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      Begin VB.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   11
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Current Protocol :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Firmware Date :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Firmware Version :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Serial Number :"
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
Attribute VB_Name = "frmATv2_RequestFirmwareVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
Dim rtncode As Variant

'Show status bar stuff
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
Me.Frame1.Visible = True
Me.ProgressBar1.Visible = True
Me.Timer1.Interval = 50
Me.Timer1.Enabled = True

Me.Refresh


rtncode = ATv2_RequestLDVModelNumber
If rtncode <> -1 Then
    Me.Label1(1).Caption = rtncode
Else
    Me.Label1(1).Caption = "N/A"
End If

rtncode = ATv2_RequestLDVSerial
If rtncode <> -1 Then
    Me.Label1(2).Caption = rtncode
Else
    Me.Label1(2).Caption = "N/A"
End If

rtncode = ATv2_RequestFirmwareVersion
If rtncode <> -1 Then
    Me.Label1(3).Caption = rtncode
Else
    Me.Label1(3).Caption = "N/A"
End If

rtncode = ATv2_RequestFirmwareDate
If rtncode <> -1 Then
    'Format date properly
    rtncode = Format(rtncode, "mm/dd/yyyy")
    
    Me.Label1(4).Caption = rtncode
Else
    Me.Label1(4).Caption = "N/A"
End If

rtncode = ATv2_RequestCurrentProtocol
If rtncode <> -1 Then
    Me.Label1(5).Caption = rtncode
Else
    Me.Label1(5).Caption = "N/A"
End If

Me.Timer1.Enabled = False
Me.Frame1.Visible = False
Me.ProgressBar1.Visible = False

Me.Refresh



End Sub

Private Sub Form_Load()
   frmATv2_RequestFirmwareVersion.Left = ((Screen.Width - 100) / 2) - (frmATv2_RequestFirmwareVersion.Width / 2)
   frmATv2_RequestFirmwareVersion.Top = ((Screen.Height - 100) / 2) - (frmATv2_RequestFirmwareVersion.Height / 2)
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
