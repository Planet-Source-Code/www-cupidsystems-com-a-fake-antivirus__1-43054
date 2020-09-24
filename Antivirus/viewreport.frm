VERSION 5.00
Begin VB.Form viewreport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "viewreport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Path"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Virus Type"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "viewreport.frx":030A
      Left            =   240
      List            =   "viewreport.frx":030C
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Type   :"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Infected :"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Following Files are Infected  By The Following Viruses."
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "viewreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

