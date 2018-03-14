VERSION 5.00
Begin VB.Form highQ 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   16113
      Caption         =   "high"
      Resize          =   0   'False
      Begin VB.CommandButton Command19 
         Caption         =   "sway:흔들리다"
         Height          =   615
         Left            =   6960
         TabIndex        =   19
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton Command18 
         Caption         =   "judge:판사"
         Height          =   615
         Left            =   4800
         TabIndex        =   18
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         Caption         =   "obey:따르다"
         Height          =   615
         Left            =   2520
         TabIndex        =   17
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         Caption         =   "institute:기관"
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command15 
         Caption         =   "aisle:통로"
         Height          =   615
         Left            =   7440
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         Caption         =   "passenger:승객"
         Height          =   615
         Left            =   5400
         TabIndex        =   14
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command13 
         Caption         =   "several:몇몇의"
         Height          =   615
         Left            =   3720
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Caption         =   "content:목차"
         Height          =   615
         Left            =   2040
         TabIndex        =   12
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "jealous:질투하는"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "slip:미끄러지다"
         Height          =   615
         Left            =   6840
         TabIndex        =   10
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "proud:자랑스러운"
         Height          =   615
         Left            =   4920
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "speech:연설"
         Height          =   615
         Left            =   3240
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "quickly:빨리"
         Height          =   615
         Left            =   1680
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "curse:악담"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "wooden:나무로 된"
         Height          =   615
         Left            =   7080
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "participate:참가하다"
         Height          =   615
         Left            =   4920
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "vow:맹세"
         Height          =   615
         Left            =   3480
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "spend:쓰다"
         Height          =   615
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "lean:기울다"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "※추후 매 업데이트시 단어가 추가됩니다"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3120
         TabIndex        =   20
         Top             =   3840
         Width           =   3615
      End
   End
End
Attribute VB_Name = "highQ"
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
high1.Show
Unload Me
End Sub

Private Sub Command10_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high10.Show
Unload Me
End Sub

Private Sub Command11_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high11.Show
Unload Me
End Sub

Private Sub Command12_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high12.Show
Unload Me
End Sub

Private Sub Command13_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high13.Show
Unload Me
End Sub

Private Sub Command14_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high14.Show
Unload Me
End Sub

Private Sub Command15_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high15.Show
Unload Me
End Sub

Private Sub Command16_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high16.Show
Unload Me
End Sub

Private Sub Command17_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high17.Show
Unload Me
End Sub

Private Sub Command18_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high18.Show
Unload Me
End Sub

Private Sub Command19_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high19.Show
Unload Me
End Sub

Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high2.Show
Unload Me
End Sub

Private Sub Command3_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high3.Show
Unload Me
End Sub

Private Sub Command4_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high4.Show
Unload Me
End Sub

Private Sub Command5_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high5.Show
Unload Me
End Sub

Private Sub Command6_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high6.Show
Unload Me
End Sub

Private Sub Command7_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high7.Show
Unload Me
End Sub

Private Sub Command8_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high8.Show
Unload Me
End Sub

Private Sub Command9_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
high9.Show
Unload Me
End Sub
