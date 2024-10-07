VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClear 
      Caption         =   "Automatic Clearing"
      Height          =   255
      Left            =   2700
      TabIndex        =   3
      Top             =   1140
      Width           =   1635
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "O&kay"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboSets 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label lblSets 
      AutoSize        =   -1  'True
      Caption         =   "Set Count:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   750
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    SetNum cboSets.ListIndex + 1
    AutoClear = CBool(Abs(Int(chkClear)))
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    On Error Resume Next
    cboSets.Clear
    For i = 1 To 150
        cboSets.AddItem i & " Sets"
    Next i
    cboSets.ListIndex = UBound(Sets) - 1
    chkClear = -Int(AutoClear)
End Sub
