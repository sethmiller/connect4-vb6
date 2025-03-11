VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5640
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   5490
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8025
      Begin VB.Timer Timer1 
         Interval        =   65
         Left            =   6480
         Top             =   4560
      End
      Begin VB.Image splashpic 
         Height          =   2880
         Left            =   0
         Picture         =   "frmSplash.frx":030A
         Top             =   4680
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Image flashpic 
         Height          =   2880
         Left            =   0
         Picture         =   "frmSplash.frx":4B8C
         Top             =   4680
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "A Product of Sethware Int'l"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Image Splash 
         Height          =   2880
         Left            =   240
         Picture         =   "frmSplash.frx":940E
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2880
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   5760
         TabIndex        =   1
         Top             =   2760
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "Connect 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   765
         Left            =   4320
         TabIndex        =   3
         Top             =   1800
         Width           =   3105
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "Welcome to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   4920
         TabIndex        =   2
         Top             =   1320
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim temp As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub Form_Load()
    
    temp = 0
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    
End Sub

Private Sub Frame1_Click()
        
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub splash_Click()
        
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub lblCompany_Click()
        
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub lblCompanyProduct_Click()
        
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub lblProductName_Click()
        
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub lblVersion_Click()
        
    Unload Me
    Load Game
    Game.Visible = True
    
End Sub

Private Sub Timer1_Timer()
    
    If temp < 4 Then
        Splash.Picture = splashpic.Picture
        lblCompany.ForeColor = &HFF0000
        temp = temp + 1
    ElseIf temp >= 4 And temp < 7 Then
        Splash.Picture = flashpic.Picture
        lblCompany.ForeColor = &H0&
        temp = temp + 1
    ElseIf temp >= 7 Then
        temp = 0
    End If
    
End Sub
