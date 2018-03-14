VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "ES2014"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.VSKIN VSKIN1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11245
      _ExtentY        =   10398
      Caption         =   "Main"
      Resize          =   0   'False
      Begin VB.CommandButton Command10 
         Caption         =   "Ãß°¡ ´Ü¾î ½ÃÇè"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   13
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFC0FF&
         Caption         =   "±â·Ïº¸±â"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   12
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0FF&
         Caption         =   "´Ü¾î Ãß°¡"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   11
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF80&
         Caption         =   "´Ü¾î °ÔÀÓ"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   10
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "¿À·ù/¹®ÀÇ"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   8
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Á¦ÀÛ"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   7
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "¾÷µ¥ÀÌÆ®È®ÀÎ"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   6
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "¼³Á¤"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   5
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "¹®Àå Å×½ºÆ®"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "´Ü¾î Å×½ºÆ®"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "´Ü¾î"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "¹®Àå"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         MaskColor       =   &H00808080&
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   6240
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "??? ´Ô ¾È³çÇÏ¼¼¿ä."
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   6000
         Width           =   4095
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   6240
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Image Image8 
         Height          =   855
         Left            =   1080
         Picture         =   "Main.frx":0802
         Top             =   3360
         Width           =   4245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   6240
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   1080
         Picture         =   "Main.frx":15A8
         Top             =   840
         Width           =   420
      End
      Begin VB.Image Image7 
         Height          =   855
         Left            =   4680
         Picture         =   "Main.frx":1A8B
         Top             =   840
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   1680
         Picture         =   "Main.frx":1FD9
         Top             =   600
         Width           =   420
      End
      Begin VB.Image Image6 
         Height          =   855
         Left            =   4080
         Picture         =   "Main.frx":2571
         Top             =   480
         Width           =   420
      End
      Begin VB.Image Image5 
         Height          =   855
         Left            =   3600
         Picture         =   "Main.frx":2A7D
         Top             =   720
         Width           =   420
      End
      Begin VB.Image Image4 
         Height          =   855
         Left            =   3000
         Picture         =   "Main.frx":2D14
         Top             =   600
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   2400
         Picture         =   "Main.frx":2FCF
         Top             =   840
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   120
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_MEMORY As Long = &H4
Private Const SND_NOSTOP As Long = &H10
Private Const SND_NOWAIT As Long = &H2000
Private Const SND_RESOURCE As Long = &H40004
Private Sub Command1_Click()
sentence.Show
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
End Sub

Private Sub Command1_GotFocus()
Command1.Caption = "sentence"
End Sub

Private Sub Command1_LostFocus()
Command1.Caption = "¹®Àå"
End Sub

Private Sub Command10_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
addwordtest.Show
Unload Me
End Sub

Private Sub Command10_GotFocus()
Command10.Caption = "add word test"
End Sub

Private Sub Command10_LostFocus()
Command10.Caption = "Ãß°¡ ´Ü¾î ½ÃÇè"
End Sub

Private Sub Command11_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
wordadd.Show
End Sub



Private Sub Command11_GotFocus()
Command11.Caption = "word add"
End Sub

Private Sub Command11_LostFocus()
Command11.Caption = "´Ü¾î Ãß°¡"

End Sub

Private Sub Command13_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
result.Show
End Sub



Private Sub Command13_GotFocus()
Command13.Caption = "show result"
End Sub

Private Sub Command13_LostFocus()
Command13.Caption = "±â·Ïº¸±â"
End Sub

Private Sub Command2_Click()
Word.Show
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
End Sub

Private Sub Command2_GotFocus()
Command2.Caption = "word"
End Sub

Private Sub Command2_LostFocus()
Command2.Caption = "´Ü¾î"
End Sub

Private Sub Command3_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
wordtest.Show

End Sub

Private Sub Command3_GotFocus()
Command3.Caption = "word test"
End Sub

Private Sub Command3_LostFocus()
Command3.Caption = "´Ü¾î Å×½ºÆ®"
End Sub
Private Sub Command4_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
sentencetest.Show
End Sub

Private Sub Command4_GotFocus()
Command4.Caption = "sentence test"
End Sub

Private Sub Command4_LostFocus()
Command4.Caption = "¹®Àå Å×½ºÆ®"
End Sub

Private Sub Command5_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
setting.Show

End Sub

Private Sub Command5_GotFocus()
Command5.Caption = "Setting"
End Sub

Private Sub Command5_LostFocus()
Command5.Caption = "¼³Á¤"
End Sub

Private Sub Command6_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
update.Show
End Sub

Private Sub Command6_GotFocus()
Command6.Caption = "Update check"
End Sub

Private Sub Command6_LostFocus()
Command6.Caption = "¾÷µ¥ÀÌÆ®È®ÀÎ"
End Sub

Private Sub Command7_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
made.Show
End Sub

Private Sub Command7_GotFocus()
Command7.Caption = "Make"
End Sub

Private Sub Command7_LostFocus()
Command7.Caption = "Á¦ÀÛ"
End Sub

Private Sub Command8_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
support.Show
End Sub


Private Sub Command8_GotFocus()
Command8.Caption = "Error/Support"
End Sub

Private Sub Command8_LostFocus()
Command8.Caption = "¿À·ù/¹®ÀÇ"
End Sub

Private Sub Command9_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
gameword.Show
Main.Hide
End Sub

Private Sub Command9_GotFocus()
Command9.Caption = "word game"
End Sub

Private Sub Command9_LostFocus()
Command9.Caption = "´Ü¾î°ÔÀÓ"
End Sub

Private Sub Form_Load()
If setting.Check3.value = 0 Then
PlaySound App.Path & "\" & "sound\int.wav", 0, SND_ASYNC
End If
Label1.Caption = Login.Text1.Text & " ´Ô ¾È³çÇÏ¼¼¿ä!"
Unload Login
End Sub

