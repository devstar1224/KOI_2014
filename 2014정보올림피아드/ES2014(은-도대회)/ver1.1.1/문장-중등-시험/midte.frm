VERSION 5.00
Begin VB.Form midte 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   Icon            =   "midte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.VSKIN VSKIN1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   20135
      _ExtentY        =   12726
      Caption         =   "����-mid"
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ü��"
         Height          =   615
         Left            =   9960
         Style           =   1  '�׷���
         TabIndex        =   20
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "��������"
         Height          =   615
         Left            =   480
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Text            =   "s"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Text            =   "j"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Text            =   "a"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '����
         Caption         =   "a fast bird. �� �׷��� ���� ���� ���� �� ���� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   4920
         TabIndex        =   18
         Top             =   4080
         Width           =   7815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '����
         Caption         =   "I have never seen"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   4080
         Width           =   5175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '����
         Caption         =   "like her mother �׳�� ��Ӵϸ� �� ��Ҵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3240
         TabIndex        =   15
         Top             =   3480
         Width           =   8055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "She looks"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         Caption         =   "go to the same store �� �� ���� ���Կ� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   2880
         Width           =   7335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "? ������ Ÿ�� �� �Ŵ� ������ ���Ŵ�?"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "Will you take the train or"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   5535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "about war ����������� ���￡ ���� �����Ѵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "Many people"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "s at 9:00 1���ô� 9�ÿ� ���۵ȴ�."
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "The first class"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
   End
End
Attribute VB_Name = "midte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell App.Path & "\" & "main.exe"
End
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
If Text1.Text = "begin" Then
Text1.Text = Text1.Text & "����"
Text1.Enabled = False
Else
Text1.Text = "����"
End If

If Text2.Text = "worry" Then
Text2.Text = Text2.Text & "����"
Text2.Enabled = False
Else
Text2.Text = "����"
End If

If Text3.Text = "fly" Then
Text3.Text = Text3.Text & "����"
Text3.Enabled = False
Else
Text3.Text = "����"
End If

If Text4.Text = "always" Then
Text4.Text = Text4.Text & "����"
Text4.Enabled = False
Else
Text4.Text = "����"
End If

If Text5.Text = "just" Then
Text5.Text = Text5.Text & "����"
Text5.Enabled = False
Else
Text5.Text = "����"
End If

If Text6.Text = "such" Then
Text6.Text = Text6.Text & "����"
Text6.Enabled = False
Else
Text6.Text = "����"
End If

End Sub
