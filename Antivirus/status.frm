VERSION 5.00
Begin VB.Form status 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.I.S  Antivirus 2002"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&About"
      Height          =   495
      Left            =   240
      MouseIcon       =   "status.frx":030A
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
      MouseIcon       =   "status.frx":0614
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
      MouseIcon       =   "status.frx":091E
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
      Height          =   495
      Left            =   240
      MouseIcon       =   "status.frx":0C28
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
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      MouseIcon       =   "status.frx":0F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      Height          =   195
      Left            =   5400
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ON"
      Height          =   195
      Left            =   5400
      TabIndex        =   17
      Top             =   3960
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Memory At First :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ON"
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   3360
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-Protect :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   13
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   0
      Picture         =   "status.frx":123C
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "28th Dec 2002."
      Height          =   195
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ON"
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   2760
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Automatically Update Antivirus :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   2730
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2nd  Feb 2002."
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Updated :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.I.S  Antivirus 2002 is protecting you from  6250  viruses."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   4980
   End
   Begin VB.Image Image1 
      Height          =   5955
      Left            =   0
      Picture         =   "status.frx":ACDE
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim end1 As Integer

 Private Sub Command1_Click()
Unload Me
status.Show
End Sub




Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub


Private Sub Command3_Click()
Unload Me
options.Show
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

 

Private Sub Label10_Click()
Label10.Visible = False
Label11.Visible = True

End Sub

Private Sub Label11_Click()
Label11.Visible = False
Label10.Visible = True

End Sub

Private Sub Label13_Click()
Label13.Visible = False
Label14.Visible = True
End Sub

Private Sub Label14_Click()
Label14.Visible = False
Label13.Visible = True
End Sub

Private Sub Label5_Click()
Label5.Visible = False
Label8.Visible = True
End Sub

Private Sub Label8_Click()
Label8.Visible = False
Label5.Visible = True
End Sub
