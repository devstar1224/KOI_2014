VERSION 5.00
Begin VB.Form elementary18 
   BorderStyle     =   0  '����
   Caption         =   "Form22"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form22"
   ScaleHeight     =   6855
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.VSKIN VSKIN1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12091
      Caption         =   "Elementary"
      Resize          =   0   'False
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         Style           =   1  '�׷���
         TabIndex        =   6
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "�ܾ�����Ű�"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         Style           =   1  '�׷���
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         Style           =   1  '�׷���
         TabIndex        =   2
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF80&
         Caption         =   "�ٷ��̵�"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         Style           =   1  '�׷���
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "bird"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   8
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   4440
         TabIndex        =   7
         Top             =   3120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "elementary18"
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
support.Show
End Sub

Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
elementary19.Show
Unload Me
End Sub

Private Sub Command3_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Main.Show
Unload Me
End Sub

Private Sub Command4_Click()
PlaySound App.Path & "\" & "word\ele\bird.wav", 0, SND_ASYNC

End Sub


Private Sub Command5_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
elementary17.Show
Unload Me

End Sub
Private Sub Command6_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
elementaryQ.Show
Unload Me

End Sub



