VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form2.frx":08CA
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "hdo01@o2.pl"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "All rights reserved"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "© Copyright 2010 by Hdo01"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   70
      Picture         =   "Form2.frx":093A
      Top             =   70
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Minesweeper Ultimate Hack"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub
