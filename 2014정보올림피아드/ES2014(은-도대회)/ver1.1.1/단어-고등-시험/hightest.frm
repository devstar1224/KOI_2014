VERSION 5.00
Begin VB.Form hightest 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   Icon            =   "hightest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.VSKIN VSKIN1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12515
      Caption         =   "high-word-test"
      Resize          =   0   'False
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Text            =   "����"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Text            =   "����"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   19
         Text            =   "����"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Text            =   "�����ϴ�"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Text            =   "������ ��"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Text            =   "�Ǵ�"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Text            =   "����"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Text            =   "����"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Text            =   "�ڶ�������"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Text            =   "�̲�������"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Text            =   "�����ϴ�"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Text            =   "����"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Text            =   "������"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Text            =   "�°�"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Text            =   "���"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Text            =   "���"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Text            =   "������"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Text            =   "�ǻ�"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   5160
         TabIndex        =   3
         Text            =   "��鸮��"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "ü��"
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
         Left            =   7080
         Style           =   1  '�׷���
         TabIndex        =   2
         Top             =   6240
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   480
         Style           =   1  '�׷���
         TabIndex        =   1
         Top             =   6240
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   2760
         X2              =   2760
         Y1              =   600
         Y2              =   5400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000B&
         X1              =   5040
         X2              =   5040
         Y1              =   600
         Y2              =   5400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000B&
         X1              =   7320
         X2              =   7320
         Y1              =   600
         Y2              =   5400
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000B&
         X1              =   480
         X2              =   9600
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   9600
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000B&
         X1              =   480
         X2              =   9600
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "�ض��� ����� ���� ���ּ���."
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   6360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "hightest"
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
Command1.Enabled = False
If Text1.Text = "lean" Then
Text1.Text = Text1.Text & " ����"
Text1.Enabled = False
Else
Text1.Text = Text1.Text & " ����"
End If

If Text2.Text = "spend" Then
Text2.Text = Text2.Text & " ����"
Text2.Enabled = False
Else
Text2.Text = Text2.Text & " ����"
End If

If Text3.Text = "vow" Then
Text3.Text = Text3.Text & " ����"
Text3.Enabled = False
Else
Text3.Text = Text3.Text & " ����"
End If

If Text4.Text = "participate" Then
Text4.Text = Text4.Text & " ����"
Text4.Enabled = False
Else
Text4.Text = Text4.Text & " ����"
End If

If Text5.Text = "wooden" Then
Text5.Text = Text5.Text & " ����"
Text5.Enabled = False
Else
Text5.Text = Text5.Text & " ����"
End If

If Text6.Text = "curse" Then
Text6.Text = Text6.Text & " ����"
Text6.Enabled = False
Else
Text6.Text = Text6.Text & " ����"
End If

If Text7.Text = "quickly" Then
Text7.Text = Text7.Text & " ����"
Text7.Enabled = False
Else
Text7.Text = Text7.Text & " ����"
End If

If Text8.Text = "speech" Then
Text8.Text = Text8.Text & " ����"
Text8.Enabled = False
Else
Text8.Text = Text8.Text & " ����"
End If

If Text9.Text = "proud" Then
Text9.Text = Text9.Text & " ����"
Text9.Enabled = False
Else
Text9.Text = Text9.Text & " ����"
End If

If Text10.Text = "slip" Then
Text10.Text = Text10.Text & " ����"
Text10.Enabled = False
Else
Text10.Text = Text10.Text & " ����"
End If

If Text11.Text = "jealous" Then
Text11.Text = Text11.Text & " ����"
Text11.Enabled = False
Else
Text11.Text = Text11.Text & " ����"
End If

If Text12.Text = "content" Then
Text12.Text = Text12.Text & " ����"
Text12.Enabled = False
Else
Text12.Text = Text12.Text & " ����"
End If

If Text13.Text = "several" Then
Text13.Text = Text13.Text & " ����"
Text13.Enabled = False
Else
Text13.Text = Text13.Text & " ����"
End If

If Text14.Text = "passenger" Then
Text14.Text = Text14.Text & " ����"
Text14.Enabled = False
Else
Text14.Text = Text14.Text & " ����"
End If

If Text15.Text = "aisle" Then
Text15.Text = Text15.Text & " ����"
Text15.Enabled = False
Else
Text15.Text = Text15.Text & " ����"
End If

If Text16.Text = "institute" Then
Text16.Text = Text16.Text & " ����"
Text16.Enabled = False
Else
Text16.Text = Text16.Text & " ����"
End If

If Text17.Text = "obey" Then
Text17.Text = Text17.Text & " ����"
Text17.Enabled = False
Else
Text17.Text = Text17.Text & " ����"
End If

If Text18.Text = "judge" Then
Text18.Text = Text18.Text & " ����"
Text18.Enabled = False
Else
Text18.Text = Text18.Text & " ����"
End If


If Text19.Text = "sway" Then
Text19.Text = Text19.Text & " ����"
Text19.Enabled = False
Else
Text19.Text = Text19.Text & " ����"
End If
End Sub

Private Sub Command2_Click()
Shell App.Path & "\" & "main.exe"
End
End Sub
