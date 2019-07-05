VERSION 5.00
Begin VB.Form index 
   BackColor       =   &H00000000&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "MiniGame for Sotaneum"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   Icon            =   "index.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   4095
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Game_play_reset 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   240
   End
   Begin VB.Timer Main_System_T 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   240
   End
   Begin VB.Label User 
      BackColor       =   &H00FF00FF&
      Caption         =   "0"
      Height          =   375
      Left            =   1920
      TabIndex        =   51
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   19
      Left            =   5640
      TabIndex        =   50
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   18
      Left            =   6600
      TabIndex        =   49
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   17
      Left            =   6480
      TabIndex        =   48
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   16
      Left            =   6240
      TabIndex        =   47
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   15
      Left            =   6000
      TabIndex        =   46
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   14
      Left            =   5760
      TabIndex        =   45
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   13
      Left            =   6720
      TabIndex        =   44
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   12
      Left            =   6120
      TabIndex        =   43
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   11
      Left            =   5880
      TabIndex        =   42
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   10
      Left            =   6360
      TabIndex        =   41
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   375
      Index           =   27
      Left            =   6240
      TabIndex        =   40
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   375
      Index           =   26
      Left            =   6720
      TabIndex        =   39
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808000&
      Caption         =   "0"
      Height          =   375
      Index           =   25
      Left            =   5760
      TabIndex        =   38
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   24
      Left            =   7440
      TabIndex        =   37
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   23
      Left            =   6120
      TabIndex        =   36
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      Height          =   375
      Index           =   22
      Left            =   8520
      TabIndex        =   35
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   375
      Index           =   21
      Left            =   7440
      TabIndex        =   34
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   375
      Index           =   20
      Left            =   6000
      TabIndex        =   33
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   375
      Index           =   19
      Left            =   6480
      TabIndex        =   32
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808000&
      Caption         =   "0"
      Height          =   375
      Index           =   18
      Left            =   5520
      TabIndex        =   31
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   17
      Left            =   7200
      TabIndex        =   30
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   16
      Left            =   5880
      TabIndex        =   29
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      Height          =   375
      Index           =   15
      Left            =   8280
      TabIndex        =   28
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   375
      Index           =   14
      Left            =   7200
      TabIndex        =   27
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   375
      Index           =   13
      Left            =   6000
      TabIndex        =   26
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   375
      Index           =   12
      Left            =   6480
      TabIndex        =   25
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808000&
      Caption         =   "0"
      Height          =   375
      Index           =   11
      Left            =   5520
      TabIndex        =   24
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   10
      Left            =   7200
      TabIndex        =   23
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   9
      Left            =   5880
      TabIndex        =   22
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      Height          =   375
      Index           =   8
      Left            =   8280
      TabIndex        =   21
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   20
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Gun_hp_bar 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   360
      TabIndex        =   19
      Top             =   6840
      Width           =   3600
   End
   Begin VB.Line Stop_bar 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   4080
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label score 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "No one"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      Height          =   375
      Index           =   5
      Left            =   8040
      TabIndex        =   16
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   4
      Left            =   5640
      TabIndex        =   15
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H0000FFFF&
      Caption         =   "0"
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   14
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808000&
      Caption         =   "0"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   13
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Unuser 
      BackColor       =   &H00808080&
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Gun_hp 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "100"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   0
      TabIndex        =   10
      Top             =   6840
      Width           =   270
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   9
      Left            =   5040
      TabIndex        =   9
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   8
      Left            =   4560
      TabIndex        =   8
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   7
      Left            =   4800
      TabIndex        =   7
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   6
      Left            =   5400
      TabIndex        =   6
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   5
      Left            =   4440
      TabIndex        =   5
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   4680
      TabIndex        =   4
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   5160
      TabIndex        =   2
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Gun 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   4320
      TabIndex        =   0
      Top             =   6840
      Width           =   135
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim gun_max, unuser_max, score_main, gun_top_max, Gun_hp_bar_m, User_hp_Max, Game_Lv, bonus As Long
'키 인식을 위한 부분입니다.
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DDown, ADown, WDown, SDown
Dim a, b, c, d, e As Long

DDown = GetAsyncKeyState(vbKeyD)
ADown = GetAsyncKeyState(vbKeyA)
WDown = GetAsyncKeyState(vbKeyW)
SDown = GetAsyncKeyState(vbKeyS)
kDown = GetAsyncKeyState(vbKeyK)

a = User.Left
b = User.Top

If DDown Then

    User.Left = a + 100
    If User.Left > index.Width - 200 Then
        User.Left = -165
    End If
       
ElseIf ADown Then

    User.Left = a - 100
    If User.Left < -165 Then
        User.Left = index.Width - 200
    End If
    
End If

If WDown Then

    User.Top = b - 110
    If User.Top < 0 Then
        User.Top = 0
    End If
    
ElseIf SDown Then

    User.Top = b + 110
    If User.Top > Stop_bar.Y1 - User.Height Then
        User.Top = Stop_bar.Y1 - User.Height
    End If

End If

If kDown Then

    For t = 0 To gun_max * 10 + 10 Step 10
    If Gun_hp.Caption > 0 Then
        If Gun_hp.Caption = t Then
        c = t / 10 - 1
        End If
    ElseIf Gun_hp.Caption = 0 Then
    
    End If
    Next t
    
    For i = c To c Step 1
    If Gun(i).Top < gun_top_max Then
    ElseIf Gun(i).Top >= gun_top_max Then
    Gun(i).Top = User.Top
    Gun(i).Left = User.Left + 120
    e = Gun_hp.Caption
    Gun_hp.Caption = e - 10
    e = Gun_hp_bar.Width
    Gun_hp_bar.Width = e - Gun_hp_bar_m / (gun_max + 1)
        
    End If
    Next i
End If

End Sub
'게임 플레이를 위한 리셋 기능 실행
Private Sub Form_Load()
Game_play_reset = True

End Sub
'게임 플레이를 위한 리셋 기능
Private Sub Game_play_reset_Timer()

'아래 사항을 수정시 실제로 값도 수정해야함.

User.Top = index.Height - 1215
User.Left = index.Width - 2265
gun_max = 19
unuser_max = 12
score_main = 0
gun_top_max = Gun_hp_bar.Top
Gun_hp.Caption = gun_max * 10 + 10
Gun_hp_bar_m = Gun_hp_bar.Width 'hp_bar 넓이 설정
User_hp_Max = 5000
bonus = 60
Game_Lv = 0

For i = 0 To unuser_max Step 1
Unuser(i).Top = Unuser(i).Top + 1 + Rnd() * i
Unuser(i).Top = -600
Unuser(i).Left = Rnd * i * 300
Unuser(i).Caption = 0
If Unuser(i).Left > index.Width Then
    Unuser(i).Left = Rnd * i * 300
End If

Next i
User.Caption = 0
Stop_bar.Y1 = index.Height - 713
Stop_bar.Y2 = index.Height - 713

For i = 0 To gun_max Step 1
Gun(i).Top = Gun_hp_bar.Top
Next i

Main_System_T.Enabled = True
Game_play_reset.Enabled = False
End Sub
'게임 플레이를 위한 기본적인 틀
Private Sub Main_System_T_Timer()
Dim Gun_top, Gun_hp_hp, Unuser_hp, score_no, e, User_hp As Long

'Gun_hp 설정
For i = 0 To gun_max
    If Gun(i).Top < gun_top_max Then
    Gun_top = Gun(i).Top
    Gun(i).Top = Gun_top - 100
    End If
Next i

'Gun_hp_bar 설정
For i = 0 To gun_max Step 1
    If Gun(i).Top <= 0 Then
    Gun(i).Top = gun_top_max
    Gun_hp_hp = Gun_hp.Caption
    Gun_hp.Caption = Gun_hp_hp + 10
    Gun(i).Left = index.Width
    e = Gun_hp_bar.Width
    Gun_hp_bar.Width = e + Gun_hp_bar_m / (gun_max + 1)
    
    End If
Next i

'Gun 과 Unuser 충돌 설정
For i = 0 To gun_max
For X = 0 To unuser_max
If Unuser(X).Top > Gun(i).Top Then
    If Unuser(X).Top < Gun(i).Top + 300 Then
        If Unuser(X).Left > Gun(i).Left - 375 Then
            If Unuser(X).Left <= Gun(i).Left + 30 Then
                Unuser_hp = Unuser(X).Caption
                Unuser(X).Caption = Unuser_hp + 7
            End If
        End If
    End If
End If
Next X
Next i

'Unuser 체력 및 데미지 설정 & 점수 표시 설정

For X = 0 To unuser_max
If Unuser(X).Caption >= 100 Then
    Unuser(X).Top = 30000
    score_no = score_main
    score_main = score_no + bonus
    score.Caption = "Lv." & Game_Lv & "  Score : " & score_main & " Hp : " & User_hp_Max
    Unuser(X).Caption = 0
    
End If
Next X

'User 게임 오버 설정

If User_hp_Max <= 0 Then
    Load index_2
    index_2.Show
    MsgBox "현재 점수 : " & score_main
    Unload index
    

End If

'Unuser 랜덤 소환 설정

For i = 0 To unuser_max Step 1
Unuser(i).Top = Unuser(i).Top + 1 + Rnd() * i
If Unuser(i).Top > Stop_bar.Y1 - 375 Then
    Unuser(i).Top = -600
    Unuser(i).Left = Rnd * i * 300
    
    If Unuser(i).Left > index.Width - 375 Then
        Unuser(i).Left = Rnd * i * 300
    End If

End If
Next i

'Unuser 와 User 와의 충돌 설정

For X = 0 To unuser_max
If Unuser(X).Top + 375 < User.Top Then
    If Unuser(X).Top - 375 > User.Top Then
        If Unuser(X).Left + 375 < User.Left Then
            If Unuser(X).Left - 375 > User.Left Then
                User_hp = User.Caption
                User.Caption = User_hp + 1
            End If
        End If
    End If
End If
If Unuser(X).Top + 375 > User.Top Then
    If Unuser(X).Top - 375 < User.Top Then
        If Unuser(X).Left + 375 > User.Left Then
            If Unuser(X).Left - 375 < User.Left Then
                User_hp = User.Caption
                User.Caption = User_hp + 1
                User_hp = User_hp_Max
                User_hp_Max = User_hp - User.Caption
                score.Caption = "Lv." & Game_Lv & "  Score : " & score_main & " Hp : " & User_hp_Max
            End If
        End If
    End If
End If
Next X

'점수에 따른 레벨 설정 & 레벨에 따른 Unuser 갯수 설정

For i = 0 To 1000000 Step 1300
If i > 0 Then
    If score_main >= i Then
        Game_Lv = i / 1300
        unuser_max = unuser_max + i
    End If
End If

If unuser_max > 26 Then
unuser_max = 27
End If

score.Caption = "Lv." & Game_Lv & "  Score : " & score_main & " Hp : " & User_hp_Max

Next i


End Sub

