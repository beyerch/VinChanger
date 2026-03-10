VERSION 5.00
Begin VB.Form frmInterfaceDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VinEditor : Select Interface Device"
   ClientHeight    =   4020
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Select Your Hardware Interface"
      Height          =   2415
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      Begin VB.OptionButton Option1 
         Caption         =   "&Tech II"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Select this if you have a GM Tech II Scantool device."
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&DHP Flash Programming Interface Cable"
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Select this option if you have the soon to be released DHP Interface"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Autotap &2.0 Cable"
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Select This option if you are using the new Autotap 2.0 Cable"
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Autotap &1.0 Cable"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Select This option if you have the original Autotap Cable"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmInterfaceDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub Form_Load()

'check current option
If intScanToolType = 0 Then intScanToolType = 1

Me.Option1.Item(intScanToolType).Value = True


End Sub

Private Sub OKButton_Click()
Dim x As Integer

'Update Current Tool item
For x = Me.Option1.LBound To Me.Option1.UBound
   If Me.Option1.Item(x).Value = True Then
      intScanToolType = x
   End If
Next x

Select Case intScanToolType
    Case 1
        strCableType = "Autotap 1.0 "
    Case 2
        strCableType = "Autotap 2.0 "
    Case 3
        strCableType = "DHP Interface "
    Case 4
        strCableType = "Tech II"
End Select

UpdateStatusBar



'Exit
Unload Me


End Sub
