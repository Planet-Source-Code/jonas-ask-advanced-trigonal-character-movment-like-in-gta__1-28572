VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2940
      Width           =   255
   End
   Begin VB.PictureBox PicManM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   5400
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox PicMan 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   4620
      Picture         =   "Form1.frx":93EA
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExitApp As Boolean 'exit application flag


Private Sub cmdExit_Click()
    ExitApp = True 'falg program for termination
End Sub

Private Sub Form_Load()
    Main.Show
    P1.X = 15
    P1.Y = 15
    MainLoop
End Sub
Sub MainLoop()
Dim NextTick As Long
    Do Until ExitApp
        DoEvents
        
        'Framelimiter
        Do Until GetTickCount > NextTick
            DoEvents
        Loop: NextTick = GetTickCount + 50
        
        DoKeys 'check for key inputs
        MoveMan 'move the man
        PaintBoard 'paint the graphics
        
    Loop
    End
End Sub
