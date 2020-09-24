VERSION 5.00
Begin VB.Form options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.I.S  Antivirus 2002"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6240
      MouseIcon       =   "options.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save Changes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      MouseIcon       =   "options.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delete The Infected File. (not recommended)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   3240
      Width           =   4935
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Automatically Repair The Infected File (recommended0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Value           =   -1  'True
      Width           =   5175
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Start Auto-Protect When Windows Starts Up (recommended)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   5655
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Enable Auto-Protect (recommended)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1680
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&About"
      Height          =   495
      Left            =   240
      MouseIcon       =   "options.frx":091E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   240
      MouseIcon       =   "options.frx":0C28
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scan For &Virusus"
      Height          =   495
      Left            =   240
      MouseIcon       =   "options.frx":0F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Options"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      MouseIcon       =   "options.frx":123C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Status"
      Height          =   495
      Left            =   240
      MouseIcon       =   "options.frx":1546
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   0
      Picture         =   "options.frx":1850
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OPTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   5955
      Left            =   0
      Picture         =   "options.frx":B2F2
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim end1 As Integer



Private Sub Check1_Click()
Command6.Enabled = True
End Sub

Private Sub Check2_Click()
Command6.Enabled = True
End Sub

 Private Sub Command1_Click()
status.Show
Me.Hide

End Sub




Private Sub Command2_Click()
Form1.Show
Me.Hide

End Sub


Private Sub Command3_Click()
options.Show
Me.Hide

End Sub



Private Sub Command5_Click()
aboutme.Show
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbWhite
Command2.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0
Command4.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub

 

 

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = vbWhite
Command1.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0
Command4.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = vbWhite
Command1.BackColor = &HE0E0E0
Command2.BackColor = &HE0E0E0
Command4.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub



Private Sub Command4_Click()
endv
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BackColor = vbWhite
Command1.BackColor = &HE0E0E0
Command2.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub



Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = vbWhite
Command4.BackColor = &HE0E0E0
Command1.BackColor = &HE0E0E0
Command2.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0

End Sub



Private Sub Command7_Click()
Unload Me
status.Show
End Sub





Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HE0E0E0
Command2.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0
Command4.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub

Public Sub endv()
end1 = MsgBox("Are You Sure You Want To Exit ?", vbQuestion + vbYesNo, "Quitting")
If end1 = vbYes Then
End
End If
End Sub

Private Sub Option1_Click()
Command6.Enabled = True
End Sub

Private Sub Option2_Click()
Command6.Enabled = True
End Sub


