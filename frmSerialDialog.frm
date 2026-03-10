VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSerialDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vin Editor : CommPort Properties"
   ClientHeight    =   4260
   ClientLeft      =   4140
   ClientTop       =   1665
   ClientWidth     =   6015
   Icon            =   "frmSerialDialog.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSettings 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   255
      TabIndex        =   1
      Top             =   570
      Width           =   5445
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   4335
         TabIndex        =   22
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Frame Frame1 
         Caption         =   "Maximum Speed"
         Height          =   870
         Left            =   180
         TabIndex        =   20
         Top             =   630
         Width           =   2340
         Begin VB.ComboBox cboSpeed 
            Enabled         =   0   'False
            Height          =   315
            Left            =   375
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.Frame fraConnection 
         Caption         =   "Connection Preferences"
         Height          =   1770
         Left            =   180
         TabIndex        =   12
         Top             =   1635
         Width           =   2325
         Begin VB.ComboBox cboStopBits 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1260
            Width           =   1140
         End
         Begin VB.ComboBox cboParity 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   810
            Width           =   1140
         End
         Begin VB.ComboBox cboDataBits 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   330
            Width           =   1140
         End
         Begin VB.Label Label5 
            Caption         =   "Stop Bits:"
            Height          =   285
            Left            =   180
            TabIndex        =   19
            Top             =   1320
            Width           =   885
         End
         Begin VB.Label Label4 
            Caption         =   "Parity:"
            Height          =   285
            Left            =   180
            TabIndex        =   18
            Top             =   855
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Data Bits:"
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   375
            Width           =   825
         End
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   300
         Left            =   4335
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   705
         Width           =   1080
      End
      Begin VB.Frame Frame7 
         Caption         =   "&Echo"
         Height          =   870
         Left            =   2595
         TabIndex        =   7
         Top             =   630
         Width           =   1590
         Begin VB.OptionButton optEcho 
            Caption         =   "Off"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   135
            MaskColor       =   &H00000000&
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optEcho 
            Caption         =   "On"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   795
            MaskColor       =   &H00000000&
            TabIndex        =   8
            Top             =   420
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Flow Control"
         Height          =   1770
         Left            =   2595
         TabIndex        =   2
         Top             =   1635
         Width           =   1620
         Begin VB.OptionButton optFlow 
            Caption         =   "None"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   6
            Top             =   345
            Width           =   855
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "Xon/Xoff"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   5
            Top             =   645
            Width           =   1095
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "RTS"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   4
            Top             =   945
            Width           =   735
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "Xon/RTS"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   180
            MaskColor       =   &H00000000&
            TabIndex        =   3
            Top             =   1245
            Width           =   1155
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         Height          =   315
         Left            =   330
         TabIndex        =   13
         Top             =   180
         Width           =   495
      End
   End
   Begin MSComctlLib.TabStrip tabSettings 
      Height          =   4065
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   7170
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSerialDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iFlow As Integer, iTempEcho As Boolean


Sub LoadPropertySettings()
Dim i As Integer, Settings As String, Offset As Integer

' Load Port Settings
For i = 1 To 16
    cboPort.AddItem "Com" & Trim$(Str$(i))
Next i

' Load Speed Settings
cboSpeed.AddItem "9600"
cboSpeed.AddItem "14400"
cboSpeed.AddItem "19200"
cboSpeed.AddItem "28800"
cboSpeed.AddItem "38400"
cboSpeed.AddItem "56000"
cboSpeed.AddItem "57600"
cboSpeed.AddItem "115200"
cboSpeed.AddItem "128000"
cboSpeed.AddItem "256000"

' Load Data Bit Settings
cboDataBits.AddItem "4"
cboDataBits.AddItem "5"
cboDataBits.AddItem "6"
cboDataBits.AddItem "7"
cboDataBits.AddItem "8"

' Load Parity Settings
cboParity.AddItem "Even"
cboParity.AddItem "Odd"
cboParity.AddItem "None"
cboParity.AddItem "Mark"
cboParity.AddItem "Space"

' Load Stop Bit Settings
cboStopBits.AddItem "1"
cboStopBits.AddItem "1.5"
cboStopBits.AddItem "2"

' Set Default Settings

Settings = frmMain.MSComm1.Settings

' In all cases the right most part of Settings will be 1 character
' except when there are 1.5 stop bits.
If InStr(Settings, ".") > 0 Then
    Offset = 2
Else
    Offset = 0
End If

cboSpeed.Text = Left$(Settings, Len(Settings) - 6 - Offset)
Select Case Mid$(Settings, Len(Settings) - 4 - Offset, 1)
Case "e"
    cboParity.ListIndex = 0
Case "m"
    cboParity.ListIndex = 1
Case "n"
    cboParity.ListIndex = 2
Case "o"
    cboParity.ListIndex = 3
Case "s"
    cboParity.ListIndex = 4
End Select

cboDataBits.Text = Mid$(Settings, Len(Settings) - 2 - Offset, 1)
cboStopBits.Text = Right$(Settings, 1 + Offset)
    
cboPort.ListIndex = frmMain.MSComm1.CommPort - 1

optFlow(frmMain.MSComm1.Handshaking).Value = True
If Echo Then
    optEcho(1).Value = True
Else
    optEcho(0).Value = True
End If

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim OldPort As Integer, ReOpen As Boolean

On Error Resume Next

Echo = iTempEcho
OldPort = frmMain.MSComm1.CommPort
NewPort = cboPort.ListIndex + 1

If NewPort <> OldPort Then                   ' If the port number changes, close the old port.
    If frmMain.MSComm1.PortOpen Then
           frmMain.MSComm1.PortOpen = False
           ReOpen = True
    End If

    frmMain.MSComm1.CommPort = NewPort          ' Set the new port number.
    
    If Err = 0 Then
        If ReOpen Then
            frmMain.MSComm1.PortOpen = True
        End If
    End If
        
    If Err Then
        MsgBox Error$, 48
        frmMain.MSComm1.CommPort = OldPort
        Exit Sub
    End If
End If


frmMain.MSComm1.Settings = Trim$(cboSpeed.Text) & "," & Left$(cboParity.Text, 1) _
    & "," & Trim$(cboDataBits.Text) & "," & Trim$(cboStopBits.Text)

If Err Then
    MsgBox Error$, 48
    Exit Sub
End If

frmMain.MSComm1.Handshaking = iFlow
If Err Then
    MsgBox Error$, 48
    Exit Sub
End If

SaveSetting App.Title, "Properties", "Settings", frmMain.MSComm1.Settings
SaveSetting App.Title, "Properties", "CommPort", frmMain.MSComm1.CommPort
SaveSetting App.Title, "Properties", "Handshaking", frmMain.MSComm1.Handshaking
SaveSetting App.Title, "Properties", "Echo", Echo

Unload Me

End Sub

Private Sub Form_Load()

' Set the form's size
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

' Size the frame to fit in the tabstrip control
fraSettings.Move tabSettings.ClientLeft, tabSettings.ClientTop

' Make sure the frame is the top most control
fraSettings.ZOrder

' Load current property settings
LoadPropertySettings

End Sub




Private Sub optEcho_Click(Index As Integer)
If Index = 1 Then
    iTempEcho = True
Else
    iTempEcho = False
End If
End Sub

Private Sub optFlow_Click(Index As Integer)
iFlow = Index
End Sub


