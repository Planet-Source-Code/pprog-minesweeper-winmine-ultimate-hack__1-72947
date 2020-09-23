VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minesweeper Ultimate Hack - by Hdo01"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   6435
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh grid"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   6000
      Width           =   1815
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   1035
      TabIndex        =   36
      Top             =   0
      Width           =   1095
      Begin VB.Image imgExtra2 
         Height          =   240
         Left            =   240
         Picture         =   "Form1.frx":1194
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgExtra1 
         Height          =   240
         Left            =   240
         Picture         =   "Form1.frx":14D6
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img8 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":1818
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img7 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":1B5A
         Stretch         =   -1  'True
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img6 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":1E9C
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img5 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":21DE
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img4 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":2520
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img3 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":2862
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img2 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":2BA4
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBomb3 
         Height          =   240
         Left            =   480
         Picture         =   "Form1.frx":2EE6
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBomb2 
         Height          =   240
         Left            =   480
         Picture         =   "Form1.frx":3228
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBlock 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   720
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgEmpty 
         Height          =   240
         Left            =   240
         Picture         =   "Form1.frx":356A
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBomb 
         Height          =   240
         Left            =   480
         Picture         =   "Form1.frx":38AC
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgNumEmpty 
         Height          =   240
         Left            =   240
         Picture         =   "Form1.frx":3BEE
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img1 
         Height          =   240
         Left            =   0
         Picture         =   "Form1.frx":3F30
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "iNFO"
      Height          =   320
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2160
      Top             =   5880
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Auto Refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Game Editor"
      Height          =   2775
      Left            =   4800
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command5 
         Caption         =   "Remove all mines"
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Put Marks on hidden mines"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "God mode"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1560
         Top             =   2400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   1800
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   1800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblMines 
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   26
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lblTime 
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   560
         TabIndex        =   25
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mines:"
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grid Editor"
      Height          =   2775
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   1560
         Picture         =   "Form1.frx":4272
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   1200
         Picture         =   "Form1.frx":45B4
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   840
         Picture         =   "Form1.frx":48F6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   480
         Picture         =   "Form1.frx":4C38
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         Picture         =   "Form1.frx":4F7A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1560
         Picture         =   "Form1.frx":52BC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1200
         Picture         =   "Form1.frx":55FE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   840
         Picture         =   "Form1.frx":5940
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   480
         Picture         =   "Form1.frx":5C82
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         Picture         =   "Form1.frx":5FC4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1560
         Picture         =   "Form1.frx":6306
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1200
         Picture         =   "Form1.frx":6648
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   840
         Picture         =   "Form1.frx":698A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         Picture         =   "Form1.frx":6CCC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdEditBlock 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Picture         =   "Form1.frx":700E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblBlock 
         Caption         =   "???"
         Height          =   195
         Left            =   1035
         TabIndex        =   9
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grid height:"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grid width:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Block index:"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   870
      End
      Begin VB.Label lblHeight 
         Caption         =   "???"
         Height          =   195
         Left            =   970
         TabIndex        =   5
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblWidth 
         Caption         =   "???"
         Height          =   195
         Left            =   970
         TabIndex        =   4
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Block type:"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblType 
         Caption         =   "???"
         Height          =   435
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   1065
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Hdo01
'hdo01@o2.pl

'http://www.hdo01.pdg.pl/

'2010-02-26 00:43 GMT+01:00, Poland

Const varGridY As Long = &H1005338    'grid height
Const varGridX As Long = &H1005334    'grid width
Const varBombs As Long = &H1005194    'bobms left
Const varTime As Long = &H100579C     'play time
Const varGrid As Long = &H1005360     'grid start here
Const varDeath As Long = &H1005000    'god mode (1 = live, 16 = death)

Dim bc As Integer                     'imgBlock index for help
Dim varBlockAddress() As Long         'all grid address here
Dim varBlockType() As String          'all grid types here

Dim varAddress As Long                'each block address here

Dim putMarks As Boolean               'extra option..
Dim removeMines As Boolean            'extra option..


Private Sub Check2_Click()
If Check2.value = 1 Then Timer2.Enabled = True  'refresh grid 1
If Check2.value = 0 Then Timer2.Enabled = False 'refresh grid 0

End Sub

Private Sub cmdEditBlock_Click(Index As Integer)

    If varAddress = 0 Then Exit Sub
    Select Case Index
    
        Case Is = 1
        WriteMemoryByte "&H" & Hex(varAddress), "&H40" 'shown empty block
        Case Is = 2
        WriteMemoryByte "&H" & Hex(varAddress), "&HF" 'hidden empty block
        Case Is = 3
        WriteMemoryByte "&H" & Hex(varAddress), "&H8F" 'hidden mine
        Case Is = 4
        WriteMemoryByte "&H" & Hex(varAddress), "&HCC" 'destroyed mine
        Case Is = 5
        WriteMemoryByte "&H" & Hex(varAddress), "&H8A" 'shown mine
        Case Is = 6
        WriteMemoryByte "&H" & Hex(varAddress), "&H41" 'number1
        Case Is = 7
        WriteMemoryByte "&H" & Hex(varAddress), "&H42" 'number2
        Case Is = 8
        WriteMemoryByte "&H" & Hex(varAddress), "&H43" 'number3
        Case Is = 9
        WriteMemoryByte "&H" & Hex(varAddress), "&H44" 'number4
        Case Is = 10
        WriteMemoryByte "&H" & Hex(varAddress), "&H45" 'number5
        Case Is = 11
        WriteMemoryByte "&H" & Hex(varAddress), "&H46" 'number6
        Case Is = 12
        WriteMemoryByte "&H" & Hex(varAddress), "&H47" 'number7
        Case Is = 13
        WriteMemoryByte "&H" & Hex(varAddress), "&H48" 'number8
        Case Is = 14
        WriteMemoryByte "&H" & Hex(varAddress), "&HE" 'marks
        Case Is = 15
        WriteMemoryByte "&H" & Hex(varAddress), "&HD" 'question marks
        End Select

        
    Call LoadGrid

End Sub

Private Sub Command1_Click()
Call LoadGrid
Call fControls
End Sub

Private Sub Command2_Click()
On Error GoTo error
Dim a As Long
a = InputBox("Enter time value:", , "0")
If Not IsNumeric(a) Then MsgBox "Wrong value!", vbExclamation: Exit Sub
WriteMemoryLong varTime, a 'write time value
Exit Sub
error:
MsgBox "Wrong value!", vbExclamation
End Sub

Private Sub Command3_Click()
On Error GoTo error
Dim a As Long
a = InputBox("Enter number of mines left:", , "0")
If Not IsNumeric(a) Then MsgBox "Wrong value!", vbExclamation: Exit Sub
WriteMemoryLong varBombs, a 'write bombs value
Exit Sub
error:
MsgBox "Wrong value!", vbExclamation
End Sub

Private Sub Command4_Click()
putMarks = True         'extra function..
Call LoadGrid
Call LoadGrid
End Sub

Private Sub Command5_Click()
removeMines = True      'extra function..
Call LoadGrid
Call LoadGrid
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Call GetWinmineHandle    'Get handle and save it to var
If WinmineHandle = 0 Then MsgBox "Run Minesweeper first.", vbExclamation: End
Call LoadGrid            'guess lol :D
Call fControls           'put all controls on their places etc.
MsgBox "If you don't see any effect, you need to refresh the Minesweeper window, just minimize and maximize them.", vbInformation
End Sub


Private Function ClearGrid()
For n = 1 To 800
    Unload imgBlock(n) 'unload all blocks from grid
Next
End Function

Private Function LoadGrid()
On Error Resume Next

Dim xxx As Integer
Dim yyy As Integer
Dim grid As Long
Dim tempBlock As Integer

'pic1.Visible = False

Call ClearGrid 'clear grid function

xxx = ReadMemory(varGridX) 'read grid width
yyy = ReadMemory(varGridY) 'read grid height

lblHeight.Caption = yyy
lblWidth.Caption = xxx

ReDim varBlockAddress(xxx * yyy) As Long 'set size of array (how many blocks in grid)
ReDim varBlockType(xxx * yyy) As String

pic1.Width = xxx * (imgBlock(0).Width) + 65   'set size of grids border
pic1.Height = yyy * (imgBlock(0).Height) + 55

' 1. This is simple grid 9x9 (beginner) in memory:
' - byte 16 is wall
' - byte 0F is empty field
' - more bytes is in form1 code..

'16 16 16 16 16 16 16 16 16 16 16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F
'16 16 16 16 16 16 16 16 16 16 16 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F 0F

' it look like this..:

'###########0000000000000000000000
'#00000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'#000000000#0000000000000000000000
'###########0000000000000000000000

bc = 0
grid = varGrid 'grid start in memory here
For Y = 1 To yyy
    For X = 1 To xxx
    
        grid = grid + 1 'read block by block from memory
        
        tempBlock = ReadMemory(grid) 'read block by block from memory
  
        bc = bc + 1
        Unload imgBlock(bc)
        Load imgBlock(bc) 'show block on grid
        imgBlock(bc).Left = (X - 1) * (imgBlock(0).Width)
        imgBlock(bc).Top = (Y - 1) * (imgBlock(0).Height)
        imgBlock(bc).Height = imgBlock(0).Height
        imgBlock(bc).Width = imgBlock(0).Width
        imgBlock(bc).Visible = True
            
        Select Case tempBlock 'recognize block type and set image
        Case Is = &H8F ' Hide Bomb
            imgBlock(bc).Picture = imgBomb3.Picture
            varBlockType(bc) = "Hidden bomb"
            If removeMines = True Then WriteMemoryByte grid, "15" 'if "function" is run, then delete all bombs
            If putMarks = True Then WriteMemoryByte grid, "142"   'if "function" is run, then mark all bombs
            
        Case Is = &H8A ' Shown Bomb
            imgBlock(bc).Picture = imgBomb.Picture
            varBlockType(bc) = "Shown bomb"
            
        Case Is = &HCC ' Destroyed Bomb
            imgBlock(bc).Picture = imgBomb2.Picture
            varBlockType(bc) = "Destroyed bomb"
            
        Case Is = &HF  ' Hide Empty
            varBlockType(bc) = "Hidden empty"
            imgBlock(bc).Picture = imgEmpty.Picture
            
        Case Is = &H40  ' Show Empty
            varBlockType(bc) = "Shown empty"
            imgBlock(bc).Picture = imgNumEmpty.Picture
            
        Case Is = &H41 ' number1
            varBlockType(bc) = "Number 1"
            imgBlock(bc).Picture = img1.Picture
        
        Case Is = &H42 ' number2
            varBlockType(bc) = "Number 2"
            imgBlock(bc).Picture = img2.Picture
        
        Case Is = &H43 ' number3
            varBlockType(bc) = "Number 3"
            imgBlock(bc).Picture = img3.Picture
        
        Case Is = &H44 ' number4
            varBlockType(bc) = "Number 4"
            imgBlock(bc).Picture = img4.Picture
        
        Case Is = &H45 ' number5
            varBlockType(bc) = "Number 5"
            imgBlock(bc).Picture = img5.Picture
        
        Case Is = &H46 ' number6
            varBlockType(bc) = "Number 6"
            imgBlock(bc).Picture = img6.Picture
         
        Case Is = &H47 ' number7
            varBlockType(bc) = "Number 7"
            imgBlock(bc).Picture = img7.Picture
        
        Case Is = &H48 ' number8
            varBlockType(bc) = "Number 8"
            imgBlock(bc).Picture = img8.Picture
            
        Case Is = &HE   ' marks (no mine)
            varBlockType(bc) = "Mark (empty)"
            imgBlock(bc).Picture = imgExtra1.Picture
            
        Case Is = &HD ' question marks (no mine)
            varBlockType(bc) = "Question Mark (empty)"
            imgBlock(bc).Picture = imgExtra2.Picture
            
        Case Is = &H8E   ' marks (mine)
            varBlockType(bc) = "Mark (bomb)"
            imgBlock(bc).Picture = imgExtra1.Picture
            
        Case Is = &H8D ' question marks (bomb)
            varBlockType(bc) = "Question Mark (bomb)"
            imgBlock(bc).Picture = imgExtra2.Picture
            
        End Select
              
        varBlockAddress(bc) = grid 'save block address in array
        
        DoEvents
        
    Next
    
    grid = grid + 1

    tempBlock = ReadMemory(grid)


'make sure that this isnt the end of 1 line
'if its end of grid (byte 16) then go next byte, and next and next and next,
    If tempBlock = 16 Then
        For n = 1 To 1000
            grid = grid + 1
            tempBlock = ReadMemory(grid)
            If tempBlock = 16 Then Exit For
        Next n
    End If
Next

removeMines = False 'turn off the function
putMarks = False    'turn off the function


'pic1.Visible = True
'and now controls height/width/left/top etc.
Call fControls
End Function


Private Function fControls()
'position/height/width/left/top etc.
Check2.Visible = True
Command6.Visible = True
Frame1.Visible = True
Frame2.Visible = True
Frame1.Left = 100 + pic1.Width
Frame2.Left = Frame1.Left + Frame2.Width + 240
Check2.Top = 180 + pic1.Height
Command1.Top = 250 + pic1.Height + Check2.Height
Command6.Top = Command1.Top + 450
Form1.Height = pic1.Height + Command1.Height + Check2.Height + Command6.Height + 900
If Form1.Height < 3335 Then Form1.Height = 3335
Form1.Width = pic1.Width + Frame1.Width + Frame2.Width + 550

End Function

Private Sub imgBlock_Click(Index As Integer)
'grid editor here
On Error Resume Next
For i = 1 To 1000
imgBlock(i).BorderStyle = 0 'remove borders from all blocks
Next i

lblBlock.Caption = Index 'show number of block on label
varAddress = varBlockAddress(Index) 'remember address of block (by index from array)
imgBlock(Index).BorderStyle = 1     'set the active frame block
lblType.Caption = varBlockType(Index) 'show what is the block type (by index from array)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
lblTime = ReadMemory(varTime) 'show current time
lblMines = ReadMemory(varBombs) 'show bombs left
If Check1.value = 1 Then WriteMemoryByte varDeath, 1 'god mode if turned on
End Sub

Private Sub Timer2_Timer()
Call LoadGrid 'refreshing (check2)
End Sub
