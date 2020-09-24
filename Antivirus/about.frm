VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cupid Antivirus 2002"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "Form1.frx":0000
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
      MouseIcon       =   "Form1.frx":030A
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
      MouseIcon       =   "Form1.frx":0614
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
      MouseIcon       =   "Form1.frx":091E
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
      MouseIcon       =   "Form1.frx":0C28
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
      MouseIcon       =   "Form1.frx":0F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5955
      Left            =   0
      Picture         =   "Form1.frx":123C
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim end1 As Integer

 

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

