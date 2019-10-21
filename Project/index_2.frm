VERSION 5.00
Begin VB.Form index_2 
   BackColor       =   &H80000012&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "MiniGame for Sotaneum"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   Icon            =   "index_2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   4095
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Setting"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Game Play"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Copyright 2012 Sotaneum All rights ....."
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   480
      TabIndex        =   5
      Top             =   6840
      Width           =   3300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "for Sotaneum"
      ForeColor       =   &H008080FF&
      Height          =   180
      Left            =   2880
      TabIndex        =   1
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Mini Game"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2880
   End
End
Attribute VB_Name = "index_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Index
Index.Show
Index.Game_play_reset = True
Unload index_2

End Sub

Private Sub Command2_Click()
MsgBox "개발중입니다."

End Sub

Private Sub Command3_Click()
Shell "explorer.exe http://duration.digimoon.net"
End

End Sub

Private Sub Command4_Click()
End

End Sub

Private Sub Form_Load()
Index.Main_System_T = False

End Sub
