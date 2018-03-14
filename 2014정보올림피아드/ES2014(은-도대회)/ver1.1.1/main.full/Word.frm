VERSION 5.00
Begin VB.Form Word 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Word"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   6015
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.VSKIN VSKIN1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      Caption         =   "Word-level"
      Resize          =   0   'False
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "°íµî"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Áßµî"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "ÃÊµî"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   5880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   5880
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "Word"
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
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
elementary1.Show
Unload Me
Unload Main
End Sub

Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
Shell App.Path & "\" & "mid.exe"
End
End Sub

Private Sub Command3_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_NOSTOP
Shell App.Path & "\" & "high.exe"
End
End Sub

