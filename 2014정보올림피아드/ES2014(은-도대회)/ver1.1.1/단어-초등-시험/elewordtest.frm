VERSION 5.00
Begin VB.Form elewordtest 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   Icon            =   "elewordtest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   10020
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
      Caption         =   "ele-word-test"
      Resize          =   0   'False
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
         TabIndex        =   42
         Top             =   6240
         Width           =   2175
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
         TabIndex        =   41
         Top             =   6240
         Width           =   2175
      End
      Begin VB.TextBox Text40 
         Height          =   375
         Left            =   5160
         TabIndex        =   40
         Text            =   "ȣ����"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text39 
         Height          =   375
         Left            =   3000
         TabIndex        =   39
         Text            =   "��ȭ��"
         Top             =   5520
         Width           =   1935
      End
      Begin VB.TextBox Text38 
         Height          =   375
         Left            =   7200
         TabIndex        =   38
         Text            =   "�ý�"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text37 
         Height          =   375
         Left            =   960
         TabIndex        =   37
         Text            =   "�ٴ�"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text36 
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         Text            =   "�ұ�"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text35 
         Height          =   375
         Left            =   7440
         TabIndex        =   35
         Text            =   "������"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text34 
         Height          =   375
         Left            =   7440
         TabIndex        =   34
         Text            =   "����"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text33 
         Height          =   375
         Left            =   7440
         TabIndex        =   33
         Text            =   "������"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text32 
         Height          =   375
         Left            =   7440
         TabIndex        =   32
         Text            =   "����"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text31 
         Height          =   375
         Left            =   7440
         TabIndex        =   31
         Text            =   "��"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text30 
         Height          =   375
         Left            =   7440
         TabIndex        =   30
         Text            =   "��"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text29 
         Height          =   375
         Left            =   7440
         TabIndex        =   29
         Text            =   "����"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text28 
         Height          =   375
         Left            =   7440
         TabIndex        =   28
         Text            =   "��"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   5160
         TabIndex        =   27
         Text            =   "����"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   5160
         TabIndex        =   26
         Text            =   "��"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Text            =   "���"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Text            =   "ȸ��"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Text            =   "��"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   5160
         TabIndex        =   22
         Text            =   "����"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   5160
         TabIndex        =   21
         Text            =   "����"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Text            =   "��"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Text            =   "����"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Text            =   "��"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Text            =   "��"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Text            =   "~�� ����"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Text            =   "���"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Text            =   "����ø"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Text            =   "����"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Text            =   "����"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Text            =   "~����"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Text            =   "����"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Text            =   "�ٽ�"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Text            =   "����"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Text            =   "~�Ŀ�"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Text            =   "�η����ϴ�"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Text            =   "�ּ�"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Text            =   "�ൿ"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Text            =   "~�� ��������"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Text            =   "~���Ͽ�"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Text            =   "�ϳ���"
         Top             =   960
         Width           =   2055
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
         TabIndex        =   43
         Top             =   6360
         Width           =   3255
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000B&
         X1              =   480
         X2              =   9600
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   9600
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000B&
         X1              =   480
         X2              =   9600
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000B&
         X1              =   7320
         X2              =   7320
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
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   2760
         X2              =   2760
         Y1              =   600
         Y2              =   5400
      End
   End
End
Attribute VB_Name = "elewordtest"
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
Command1.Enabled = False
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
If Text1.Text = "a" Then
Text1.Text = Text1.Text & " ����"
Text1.Enabled = False
Else
Text1.Text = Text1.Text & " ����"
End If

If Text2.Text = "about" Then
Text2.Text = Text2.Text & " ����"
Text2.Enabled = False
Else
Text2.Text = Text2.Text & " ����"
End If

If Text3.Text = "across" Then
Text3.Text = Text3.Text & " ����"
Text3.Enabled = False
Else
Text3.Text = Text3.Text & " ����"
End If

If Text4.Text = "act" Then
Text4.Text = Text4.Text & " ����"
Text4.Enabled = False
Else
Text4.Text = Text4.Text & " ����"
End If

If Text5.Text = "address" Then
Text5.Text = Text5.Text & " ����"
Text5.Enabled = False
Else
Text5.Text = Text5.Text & " ����"
End If

If Text6.Text = "afraid" Then
Text6.Text = Text6.Text & " ����"
Text6.Enabled = False
Else
Text6.Text = Text6.Text & " ����"
End If

If Text7.Text = "after" Then
Text7.Text = Text7.Text & " ����"
Text7.Enabled = False
Else
Text7.Text = Text7.Text & " ����"
End If

If Text8.Text = "afternoon" Then
Text8.Text = Text8.Text & " ����"
Text8.Enabled = False
Else
Text8.Text = Text8.Text & " ����"
End If

If Text9.Text = "again" Then
Text9.Text = Text9.Text & " ����"
Text9.Enabled = False
Else
Text9.Text = Text9.Text & " ����"
End If

If Text10.Text = "age" Then
Text10.Text = Text10.Text & " ����"
Text10.Enabled = False
Else
Text10.Text = Text10.Text & " ����"
End If

If Text11.Text = "ago" Then
Text11.Text = Text11.Text & " ����"
Text11.Enabled = False
Else
Text11.Text = Text11.Text & " ����"
End If

If Text12.Text = "air" Then
Text12.Text = Text12.Text & " ����"
Text12.Enabled = False
Else
Text12.Text = Text12.Text & " ����"
End If

If Text13.Text = "airport" Then
Text13.Text = Text13.Text & " ����"
Text13.Enabled = False
Else
Text13.Text = Text13.Text & " ����"
End If

If Text14.Text = "album" Then
Text14.Text = Text14.Text & " ����"
Text14.Enabled = False
Else
Text14.Text = Text14.Text & " ����"
End If

If Text15.Text = "all" Then
Text15.Text = Text15.Text & " ����"
Text15.Enabled = False
Else
Text15.Text = Text15.Text & " ����"
End If

If Text16.Text = "along" Then
Text16.Text = Text16.Text & " ����"
Text16.Enabled = False
Else
Text16.Text = Text16.Text & " ����"
End If

If Text17.Text = "bear" Then
Text17.Text = Text17.Text & " ����"
Text17.Enabled = False
Else
Text17.Text = Text17.Text & " ����"
End If

If Text18.Text = "bird" Then
Text18.Text = Text18.Text & " ����"
Text18.Enabled = False
Else
Text18.Text = Text18.Text & " ����"
End If


If Text19.Text = "button" Then
Text19.Text = Text19.Text & " ����"
Text19.Enabled = False
Else
Text19.Text = Text19.Text & " ����"
End If


If Text20.Text = "door" Then
Text20.Text = Text20.Text & " ����"
Text20.Enabled = False
Else
Text20.Text = Text20.Text & " ����"
End If


If Text21.Text = "earth" Then
Text21.Text = Text21.Text & " ����"
Text21.Enabled = False
Else
Text21.Text = Text21.Text & " ����"
End If


If Text22.Text = "family" Then
Text22.Text = Text22.Text & " ����"
Text22.Enabled = False
Else
Text22.Text = Text22.Text & " ����"
End If


If Text23.Text = "flower" Then
Text23.Text = Text23.Text & " ����"
Text23.Enabled = False
Else
Text23.Text = Text23.Text & " ����"
End If


If Text24.Text = "gray" Then
Text24.Text = Text24.Text & " ����"
Text24.Enabled = False
Else
Text24.Text = Text24.Text & " ����"
End If


If Text25.Text = "green" Then
Text25.Text = Text25.Text & " ����"
Text25.Enabled = False
Else
Text25.Text = Text25.Text & " ����"
End If


If Text26.Text = "hand" Then
Text26.Text = Text26.Text & " ����"
Text26.Enabled = False
Else
Text26.Text = Text26.Text & " ����"
End If


If Text27.Text = "hat" Then
Text27.Text = Text27.Text & " ����"
Text27.Enabled = False
Else
Text27.Text = Text27.Text & " ����"
End If


If Text28.Text = "king" Then
Text28.Text = Text28.Text & " ����"
Text28.Enabled = False
Else
Text28.Text = Text28.Text & " ����"
End If


If Text29.Text = "key" Then
Text29.Text = Text29.Text & " ����"
Text29.Enabled = False
Else
Text29.Text = Text29.Text & " ����"
End If


If Text30.Text = "mountain" Then
Text30.Text = Text30.Text & " ����"
Text30.Enabled = False
Else
Text30.Text = Text30.Text & " ����"
End If

If Text31.Text = "nose" Then
Text31.Text = Text31.Text & " ����"
Text31.Enabled = False
Else
Text31.Text = Text31.Text & " ����"
End If

If Text32.Text = "number" Then
Text32.Text = Text32.Text & " ����"
Text32.Enabled = False
Else
Text32.Text = Text32.Text & " ����"
End If

If Text33.Text = "orange" Then
Text33.Text = Text33.Text & " ����"
Text33.Enabled = False
Else
Text33.Text = Text33.Text & " ����"
End If

If Text34.Text = "park" Then
Text34.Text = Text34.Text & " ����"
Text34.Enabled = False
Else
Text34.Text = Text34.Text & " ����"
End If

If Text35.Text = "rainbow" Then
Text35.Text = Text35.Text & " ����"
Text35.Enabled = False
Else
Text35.Text = Text35.Text & " ����"
End If

If Text36.Text = "salt" Then
Text36.Text = Text36.Text & " ����"
Text36.Enabled = False
Else
Text36.Text = Text36.Text & " ����"
End If

If Text37.Text = "sea" Then
Text37.Text = Text37.Text & " ����"
Text37.Enabled = False
Else
Text37.Text = Text37.Text & " ����"
End If

If Text38.Text = "taxi" Then
Text38.Text = Text38.Text & " ����"
Text38.Enabled = False
Else
Text38.Text = Text38.Text & " ����"
End If

If Text39.Text = "telephone" Then
Text39.Text = Text39.Text & " ����"
Text39.Enabled = False
Else
Text39.Text = Text39.Text & " ����"
End If

If Text40.Text = "tiger" Then
Text40.Text = Text40.Text & " ����"
Text40.Enabled = False
Else
Text40.Text = Text40.Text & " ����"
End If


End Sub

Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Shell App.Path & "\" & "main.exe"
End
End Sub
