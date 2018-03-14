VERSION 5.00
Begin VB.Form sentence 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "sentence"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.VSKIN VSKIN1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3413
      Caption         =   "sentence-level"
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
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
         Left            =   3120
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
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
         Left            =   960
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   5400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   120
         X2              =   5400
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "sentence"
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
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Shell App.Path & "\" & "elest.exe"
End
End Sub

Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Shell App.Path & "\" & "midst.exe"
End
End Sub

Private Sub Command3_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Shell App.Path & "\" & "highst.exe"
End
End Sub

