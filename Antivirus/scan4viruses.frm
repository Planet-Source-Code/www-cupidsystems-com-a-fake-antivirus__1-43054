VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.I.S  Antivirus 2002"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "SCAN4V~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5040
      TabIndex        =   13
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   7200
      Top             =   840
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop Scan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Start Scan"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   4560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&About"
      Height          =   495
      Left            =   240
      MouseIcon       =   "SCAN4V~1.frx":030A
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
      MouseIcon       =   "SCAN4V~1.frx":0614
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
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      MouseIcon       =   "SCAN4V~1.frx":091E
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
      MouseIcon       =   "SCAN4V~1.frx":0C28
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
      MouseIcon       =   "SCAN4V~1.frx":0F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File (s) Present In The Folder/Drive"
      Height          =   195
      Left            =   5040
      TabIndex        =   16
      Top             =   2760
      Width           =   2475
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder (s) Present In The Drive"
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Top             =   2760
      Width           =   2160
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   0
      Picture         =   "SCAN4V~1.frx":123C
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6000
      MouseIcon       =   "SCAN4V~1.frx":ACDE
      MousePointer    =   99  'Custom
      Picture         =   "SCAN4V~1.frx":AFE8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Done :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   12
      Top             =   4080
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3360
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select The Drive You Want To Scan:-"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   5955
      Left            =   0
      Picture         =   "SCAN4V~1.frx":BA47
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

 

Private Sub Command1_Click()
status.Show
Me.Hide

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbWhite
Command2.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0
Command4.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub

 

 

Private Sub Command2_Click()
Form1.Show
Me.Hide

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = vbWhite
Command1.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0
Command4.BackColor = &HE0E0E0
Command5.BackColor = &HE0E0E0
End Sub

 

Private Sub Command3_Click()
options.Show
Me.Hide

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



Private Sub Command5_Click()
aboutme.Show
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = vbWhite
Command4.BackColor = &HE0E0E0
Command1.BackColor = &HE0E0E0
Command2.BackColor = &HE0E0E0
Command3.BackColor = &HE0E0E0

End Sub

Private Sub Command6_Click()
Label3.Visible = False
Image2.Visible = False
Timer1.Enabled = True
Command7.Enabled = True
Command6.Enabled = False
ProgressBar1.Value = 0
Label2.Caption = ProgressBar1.Value
End Sub

Private Sub Command7_Click()
Timer1.Enabled = False
Command7.Enabled = False
Command6.Enabled = True
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
On Error GoTo fault
ProgressBar1.Value = 0
Label2.Caption = ProgressBar1.Value
Label3.Visible = False
Image2.Visible = False
If Drive1.Drive = "a:" Then
Timer1.Interval = 500
Else
Timer1.Interval = 2500
End If
Dir1.Path = Drive1
File1.Path = Drive1
Exit Sub
fault:
MsgBox Err.Description & ". Please Insert Disk In The Drive.", 16 + vbOKOnly, "Error"
Drive1.Drive = "C:"
Resume Next


End Sub

Private Sub Form_Load()
Drive1.Drive = "c:"
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

Private Sub Image2_Click()
viewreport.Show
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label2.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = 100 Then
Timer1.Enabled = False
Label3.Visible = True
Image2.Visible = True
End If
End Sub
