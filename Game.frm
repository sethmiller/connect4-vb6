VERSION 5.00
Begin VB.Form Game 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect 4"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Flash 
      Interval        =   50
      Left            =   240
      Top             =   600
   End
   Begin VB.Timer Delay 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   240
      Top             =   4800
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   8
      Left            =   7440
      Tag             =   "1"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image flashbox 
      Height          =   720
      Left            =   120
      Picture         =   "Game.frx":030A
      Top             =   1800
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   7
      Left            =   6720
      Tag             =   "1"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   6
      Left            =   6000
      Tag             =   "2"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   5
      Left            =   5280
      Tag             =   "0"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   4
      Left            =   4560
      Tag             =   "1"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   3
      Left            =   3840
      Tag             =   "2"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   62
      Left            =   7320
      Picture         =   "Game.frx":080C
      Tag             =   "18"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   61
      Left            =   6600
      Picture         =   "Game.frx":0D0E
      Tag             =   "19"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   60
      Left            =   5880
      Picture         =   "Game.frx":1210
      Tag             =   "20"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   59
      Left            =   5160
      Picture         =   "Game.frx":1712
      Tag             =   "18"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   58
      Left            =   4440
      Picture         =   "Game.frx":1C14
      Tag             =   "19"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   57
      Left            =   3720
      Picture         =   "Game.frx":2116
      Tag             =   "20"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   56
      Left            =   3000
      Picture         =   "Game.frx":2618
      Tag             =   "18"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   55
      Left            =   2280
      Picture         =   "Game.frx":2B1A
      Tag             =   "19"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   54
      Left            =   1560
      Picture         =   "Game.frx":301C
      Tag             =   "20"
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   53
      Left            =   7320
      Picture         =   "Game.frx":351E
      Tag             =   "18"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   52
      Left            =   6600
      Picture         =   "Game.frx":3A20
      Tag             =   "19"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   51
      Left            =   5880
      Picture         =   "Game.frx":3F22
      Tag             =   "20"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   50
      Left            =   5160
      Picture         =   "Game.frx":4424
      Tag             =   "18"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   49
      Left            =   4440
      Picture         =   "Game.frx":4926
      Tag             =   "19"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   48
      Left            =   3720
      Picture         =   "Game.frx":4E28
      Tag             =   "20"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   47
      Left            =   3000
      Picture         =   "Game.frx":532A
      Tag             =   "9"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   46
      Left            =   2280
      Picture         =   "Game.frx":582C
      Tag             =   "10"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   45
      Left            =   1560
      Picture         =   "Game.frx":5D2E
      Tag             =   "11"
      Top             =   5880
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   44
      Left            =   7320
      Picture         =   "Game.frx":6230
      Tag             =   "12"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   43
      Left            =   6600
      Picture         =   "Game.frx":6732
      Tag             =   "13"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   42
      Left            =   5880
      Picture         =   "Game.frx":6C34
      Tag             =   "14"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   41
      Left            =   5160
      Picture         =   "Game.frx":7136
      Tag             =   "15"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   40
      Left            =   4440
      Picture         =   "Game.frx":7638
      Tag             =   "16"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   39
      Left            =   3720
      Picture         =   "Game.frx":7B3A
      Tag             =   "17"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   38
      Left            =   3000
      Picture         =   "Game.frx":803C
      Tag             =   "18"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   37
      Left            =   2280
      Picture         =   "Game.frx":853E
      Tag             =   "19"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   36
      Left            =   1560
      Picture         =   "Game.frx":8A40
      Tag             =   "20"
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   35
      Left            =   7320
      Picture         =   "Game.frx":8F42
      Tag             =   "18"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   34
      Left            =   6600
      Picture         =   "Game.frx":9444
      Tag             =   "19"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   33
      Left            =   5880
      Picture         =   "Game.frx":9946
      Tag             =   "20"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   32
      Left            =   5160
      Picture         =   "Game.frx":9E48
      Tag             =   "18"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   31
      Left            =   4440
      Picture         =   "Game.frx":A34A
      Tag             =   "19"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   30
      Left            =   3720
      Picture         =   "Game.frx":A84C
      Tag             =   "20"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   29
      Left            =   3000
      Picture         =   "Game.frx":AD4E
      Tag             =   "18"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   28
      Left            =   2280
      Picture         =   "Game.frx":B250
      Tag             =   "19"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   27
      Left            =   1560
      Picture         =   "Game.frx":B752
      Tag             =   "20"
      Top             =   4440
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   26
      Left            =   7320
      Picture         =   "Game.frx":BC54
      Tag             =   "18"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   25
      Left            =   6600
      Picture         =   "Game.frx":C156
      Tag             =   "19"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   24
      Left            =   5880
      Picture         =   "Game.frx":C658
      Tag             =   "20"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   23
      Left            =   5160
      Picture         =   "Game.frx":CB5A
      Tag             =   "18"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   22
      Left            =   4440
      Picture         =   "Game.frx":D05C
      Tag             =   "19"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   21
      Left            =   3720
      Picture         =   "Game.frx":D55E
      Tag             =   "20"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   7
      Left            =   6600
      Picture         =   "Game.frx":DA60
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   6
      Left            =   5880
      Picture         =   "Game.frx":DBD2
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   5
      Left            =   5160
      Picture         =   "Game.frx":DD44
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   4
      Left            =   4440
      Picture         =   "Game.frx":DEB6
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   3
      Left            =   3720
      Picture         =   "Game.frx":E028
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   2
      Left            =   3000
      Picture         =   "Game.frx":E19A
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image pushbox 
      Height          =   720
      Left            =   120
      Picture         =   "Game.frx":E30C
      Top             =   2520
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   2
      Left            =   3120
      Tag             =   "2"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   1
      Left            =   2400
      Tag             =   "1"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image peak 
      Height          =   495
      Index           =   0
      Left            =   1680
      Tag             =   "0"
      Top             =   1665
      Width           =   615
   End
   Begin VB.Image Image13 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":E80E
      Top             =   5880
      Width           =   165
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   20
      Left            =   3000
      Picture         =   "Game.frx":EA10
      Tag             =   "20"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   19
      Left            =   2280
      Picture         =   "Game.frx":EF12
      Tag             =   "19"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   18
      Left            =   1560
      Picture         =   "Game.frx":F414
      Tag             =   "18"
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Image12 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":F916
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   17
      Left            =   7320
      Picture         =   "Game.frx":FB18
      Tag             =   "17"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   16
      Left            =   6600
      Picture         =   "Game.frx":1001A
      Tag             =   "16"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   15
      Left            =   5880
      Picture         =   "Game.frx":1051C
      Tag             =   "15"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image Image11 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":10A1E
      Top             =   4440
      Width           =   165
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   14
      Left            =   5160
      Picture         =   "Game.frx":10C20
      Tag             =   "14"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   13
      Left            =   4440
      Picture         =   "Game.frx":11122
      Tag             =   "13"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   12
      Left            =   3720
      Picture         =   "Game.frx":11624
      Tag             =   "12"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image Image7 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":11B26
      Top             =   6600
      Width           =   165
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   11
      Left            =   3000
      Picture         =   "Game.frx":11D28
      Tag             =   "11"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   10
      Left            =   2280
      Picture         =   "Game.frx":1222A
      Tag             =   "10"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   9
      Left            =   1560
      Picture         =   "Game.frx":1272C
      Tag             =   "9"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image corner 
      Height          =   150
      Left            =   8040
      Picture         =   "Game.frx":12C2E
      Top             =   2160
      Width           =   165
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   1
      Left            =   2280
      Picture         =   "Game.frx":12D00
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   8
      Left            =   7320
      Picture         =   "Game.frx":12E72
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image topmiddle 
      Height          =   150
      Index           =   0
      Left            =   1560
      Picture         =   "Game.frx":12FE4
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":13156
      Top             =   3720
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":13358
      Top             =   3000
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   8040
      Picture         =   "Game.frx":1355A
      Top             =   2280
      Width           =   165
   End
   Begin VB.Image emptybox 
      Height          =   720
      Left            =   120
      Picture         =   "Game.frx":1375C
      Top             =   3240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label winner 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image blank 
      Height          =   480
      Left            =   240
      Picture         =   "Game.frx":13C5E
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image arrow 
      Height          =   480
      Left            =   240
      Picture         =   "Game.frx":13F68
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image no 
      Height          =   480
      Left            =   240
      Picture         =   "Game.frx":14272
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image stopsign 
      Height          =   480
      Left            =   240
      Picture         =   "Game.frx":1457C
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   4
      Left            =   4440
      Picture         =   "Game.frx":14886
      Tag             =   "4"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   5
      Left            =   5160
      Picture         =   "Game.frx":14D88
      Tag             =   "5"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   6
      Left            =   5880
      Picture         =   "Game.frx":1528A
      Tag             =   "6"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   7
      Left            =   6600
      Picture         =   "Game.frx":1578C
      Tag             =   "7"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   2
      Left            =   3000
      Picture         =   "Game.frx":15C8E
      Tag             =   "2"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   8
      Left            =   7320
      Picture         =   "Game.frx":16190
      Tag             =   "8"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   1
      Left            =   2280
      Picture         =   "Game.frx":16692
      Tag             =   "1"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   3
      Left            =   3720
      Picture         =   "Game.frx":16B94
      Tag             =   "3"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image blueball 
      DragIcon        =   "Game.frx":17096
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   4560
      Picture         =   "Game.frx":173A0
      Tag             =   "Blue"
      Top             =   960
      Width           =   480
   End
   Begin VB.Image bluebox 
      Height          =   720
      Left            =   120
      Picture         =   "Game.frx":176AA
      Top             =   3960
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image redbox 
      Height          =   720
      Left            =   120
      Picture         =   "Game.frx":17BAC
      Top             =   1080
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image box 
      Height          =   720
      Index           =   0
      Left            =   1560
      Picture         =   "Game.frx":180AE
      Tag             =   "0"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image redball 
      DragIcon        =   "Game.frx":185B0
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   4560
      Picture         =   "Game.frx":188BA
      Tag             =   "Red"
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim howfar As Integer
Dim whichColumn As Integer
Dim temp As Integer
Dim flashtemp As Integer
Private Sub box_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    peak_DragDrop Index Mod 9, Source, X!, Y!
    
End Sub

Private Sub box_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    peak_DragOver Index Mod 9, Source, X!, Y!, State%

End Sub

Private Sub box_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Delay.Enabled = False
    box(Index).Picture = pushbox.Picture
    
End Sub

Private Sub box_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    box(Index).Picture = emptybox.Picture
    'Delay.Enabled = True
    
End Sub
Private Sub cmdNew_Click()
    
    blueball.Visible = True
    redball.Visible = False
    over = False
    
    For X = 0 To 8
        peak(X).Picture = blank.Picture
    Next X
    
    For X = 0 To 6
        For Y = 0 To 9
            board(X, Y) = 0
        Next Y
    Next X
    
    winner.Caption = " "
    For X = 0 To 62
        box(X).Picture = emptybox.Picture
    Next X
        
        
End Sub
Private Sub Delay_Timer()
        
        If temp > 0 Then
            box(((temp - 1) * 9) + whichColumn).Picture = emptybox.Picture
        End If
        
        If COLOR = "Blue" Then
            box((temp * 9) + whichColumn).Picture = bluebox.Picture
            temp = temp + 1
            Delay.Interval = Delay.Interval / 2
        End If
        
        If COLOR = "Red" Then
            box((temp * 9) + whichColumn).Picture = redbox.Picture
            temp = temp + 1
            Delay.Interval = Delay.Interval / 2
        End If
        
        If temp = howfar + 1 Then
            Delay.Enabled = False
            Delay.Interval = 125
            temp = 0
            Exit Sub
        End If
        
End Sub

Private Sub Flash_Timer()
        
    If over Then
        If direction = 1 Then
            If COLOR = "Blue" And flashtemp > 4 Then
                box((9 * ecks) + why).Picture = flashbox.Picture
                box((9 * ecks) + (why + 1)).Picture = flashbox.Picture
                box((9 * ecks) + (why + 2)).Picture = flashbox.Picture
                box((9 * ecks) + (why + 3)).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Blue" And flashtemp <= 4 Then
                box((9 * ecks) + why).Picture = bluebox.Picture
                box((9 * ecks) + (why + 1)).Picture = bluebox.Picture
                box((9 * ecks) + (why + 2)).Picture = bluebox.Picture
                box((9 * ecks) + (why + 3)).Picture = bluebox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp > 4 Then
                box((9 * ecks) + why).Picture = flashbox.Picture
                box((9 * ecks) + (why + 1)).Picture = flashbox.Picture
                box((9 * ecks) + (why + 2)).Picture = flashbox.Picture
                box((9 * ecks) + (why + 3)).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp <= 4 Then
                box((9 * ecks) + why).Picture = redbox.Picture
                box((9 * ecks) + (why + 1)).Picture = redbox.Picture
                box((9 * ecks) + (why + 2)).Picture = redbox.Picture
                box((9 * ecks) + (why + 3)).Picture = redbox.Picture
                flashtemp = flashtemp + 1
            End If
        End If
        If direction = 2 Then
            If COLOR = "Blue" And flashtemp > 4 Then
                box((9 * ecks) + why).Picture = flashbox.Picture
                box((9 * (ecks + 1)) + why).Picture = flashbox.Picture
                box((9 * (ecks + 2)) + why).Picture = flashbox.Picture
                box((9 * (ecks + 3)) + why).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Blue" And flashtemp <= 4 Then
                box((9 * ecks) + why).Picture = bluebox.Picture
                box((9 * (ecks + 1)) + why).Picture = bluebox.Picture
                box((9 * (ecks + 2)) + why).Picture = bluebox.Picture
                box((9 * (ecks + 3)) + why).Picture = bluebox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp > 4 Then
                box((9 * ecks) + why).Picture = flashbox.Picture
                box((9 * (ecks + 1)) + why).Picture = flashbox.Picture
                box((9 * (ecks + 2)) + why).Picture = flashbox.Picture
                box((9 * (ecks + 3)) + why).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp <= 4 Then
                box((9 * ecks) + why).Picture = redbox.Picture
                box((9 * (ecks + 1)) + why).Picture = redbox.Picture
                box((9 * (ecks + 2)) + why).Picture = redbox.Picture
                box((9 * (ecks + 3)) + why).Picture = redbox.Picture
                flashtemp = flashtemp + 1
            End If
        End If
        If direction = 3 Then
            If COLOR = "Blue" And flashtemp > 4 Then
                box((9 * ecks) + why).Picture = flashbox.Picture
                box((9 * (ecks + 1)) + (why + 1)).Picture = flashbox.Picture
                box((9 * (ecks + 2)) + (why + 2)).Picture = flashbox.Picture
                box((9 * (ecks + 3)) + (why + 3)).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Blue" And flashtemp <= 4 Then
                box((9 * ecks) + why).Picture = bluebox.Picture
                box((9 * (ecks + 1)) + (why + 1)).Picture = bluebox.Picture
                box((9 * (ecks + 2)) + (why + 2)).Picture = bluebox.Picture
                box((9 * (ecks + 3)) + (why + 3)).Picture = bluebox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp > 4 Then
                box((9 * ecks) + why).Picture = flashbox.Picture
                box((9 * (ecks + 1)) + (why + 1)).Picture = flashbox.Picture
                box((9 * (ecks + 2)) + (why + 2)).Picture = flashbox.Picture
                box((9 * (ecks + 3)) + (why + 3)).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp <= 4 Then
                box((9 * ecks) + why).Picture = redbox.Picture
                box((9 * (ecks + 1)) + (why + 1)).Picture = redbox.Picture
                box((9 * (ecks + 2)) + (why + 2)).Picture = redbox.Picture
                box((9 * (ecks + 3)) + (why + 3)).Picture = redbox.Picture
                flashtemp = flashtemp + 1
            End If
        End If
        If direction = 4 Then
            If COLOR = "Blue" And flashtemp > 4 Then
                box((9 * ecks) + (why + 3)).Picture = flashbox.Picture
                box((9 * (ecks + 1)) + (why + 2)).Picture = flashbox.Picture
                box((9 * (ecks + 2)) + (why + 1)).Picture = flashbox.Picture
                box((9 * (ecks + 3)) + why).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Blue" And flashtemp <= 4 Then
                box((9 * ecks) + (why + 3)).Picture = bluebox.Picture
                box((9 * (ecks + 1)) + (why + 2)).Picture = bluebox.Picture
                box((9 * (ecks + 2)) + (why + 1)).Picture = bluebox.Picture
                box((9 * (ecks + 3)) + why).Picture = bluebox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp > 4 Then
                box((9 * ecks) + (why + 3)).Picture = flashbox.Picture
                box((9 * (ecks + 1)) + (why + 2)).Picture = flashbox.Picture
                box((9 * (ecks + 2)) + (why + 1)).Picture = flashbox.Picture
                box((9 * (ecks + 3)) + why).Picture = flashbox.Picture
                flashtemp = flashtemp + 1
            End If
            If COLOR = "Red" And flashtemp <= 4 Then
                box((9 * ecks) + (why + 3)).Picture = redbox.Picture
                box((9 * (ecks + 1)) + (why + 2)).Picture = redbox.Picture
                box((9 * (ecks + 2)) + (why + 1)).Picture = redbox.Picture
                box((9 * (ecks + 3)) + why).Picture = redbox.Picture
                flashtemp = flashtemp + 1
            End If
        End If
        If flashtemp >= 9 Then
            flashtemp = 0
        End If
    End If
    
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    For X = 0 To 8
        peak(X).Picture = blank.Picture
    Next X
        
End Sub
Sub renew()

    For X = 0 To 6
        For Y = 0 To 8
            If board(X, Y) = 1 Then
                box((9 * X) + Y).Picture = bluebox.Picture
            ElseIf board(X, Y) = 2 Then
                box((9 * X) + Y).Picture = redbox.Picture
            ElseIf board(X, Y) = 0 Then
                box((9 * X) + Y).Picture = emptybox.Picture
            End If
        Next Y
    Next X
    'winner.Caption = check()

End Sub
Sub Drop(column As Integer, Source As Control)
    
     For X% = 6 To 0 Step -1
        If board(X%, column) = 0 Then
            If Source.tag = "Blue" Then
                board(X%, column) = 1
                'box(9 * X% + column).Picture = bluebox.Picture
                
                COLOR = "Blue"
                whichColumn = column
                howfar = X%
                Delay.Enabled = True
                
                peak(column).Picture = blank.Picture
                blueball.Visible = False
                redball.Visible = True
                winner.Caption = check("Blue")
                Exit Sub
            End If
            If Source.tag = "Red" Then
                board(X%, column) = 2
                'box(9 * X% + column).Picture = redbox.Picture
                
                COLOR = "Red"
                whichColumn = column
                howfar = X%
                Delay.Enabled = True
                
                peak(column).Picture = blank.Picture
                redball.Visible = False
                blueball.Visible = True
                winner.Caption = check("Red")
                Exit Sub
            End If
        End If
    Next X%
    
    peak(column).Picture = blank.Picture
    
    
End Sub

Private Sub peak_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    
    Drop Index, Source
    
End Sub

Private Sub peak_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    
    If over Then
        For X = 0 To 8
            If X <> Index Then
                peak(X).Picture = blank.Picture
            End If
        peak(Index).Picture = stopsign.Picture
        Next X
    Else
        For X = 0 To 8
            If X <> Index Then
                peak(X).Picture = blank.Picture
            End If
        Next X
    peak(Index).Picture = arrow.Picture
    End If
    
End Sub

Private Sub topmiddle_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    peak_DragDrop Index Mod 9, Source, X!, Y!


End Sub

Private Sub topmiddle_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    peak_DragOver Index Mod 9, Source, X!, Y!, State%

End Sub
