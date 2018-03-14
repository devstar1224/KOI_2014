VERSION 5.00
Begin VB.Form mid28 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Form11"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form11"
   ScaleHeight     =   6855
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.VSKIN VSKIN1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12091
      Caption         =   "Mid"
      Resize          =   0   'False
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF80&
         Caption         =   "¹Ù·ÎÀÌµ¿"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "¹ßÀ½"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   4
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "¸ÞÀÎÀ¸·Î"
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
         Left            =   3960
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   3
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "´ÙÀ½"
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
         Left            =   7680
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   2
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FF8080&
         Caption         =   "ÀÌÀü"
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
         Left            =   600
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   1
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Áþ´Ù"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   21.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   4320
         TabIndex        =   7
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "build"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   3720
         TabIndex        =   6
         Top             =   1920
         Width           =   4695
      End
   End
End
Attribute VB_Name = "mid28"
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


Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
mid29.Show
Unload Me
End Sub

Private Sub Command3_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Shell App.Path & "\" & "main.exe"
End
End Sub

Private Sub Command4_Click()
PlaySound App.Path & "\" & "word\mid\build.wav", 0, SND_ASYNC

End Sub

Private Sub Command6_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
midQ.Show
Unload Me

End Sub
Private Sub Command5_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
mid27.Show
Unload Me

End Sub


