VERSION 5.00
Begin VB.Form midst 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   Icon            =   "midst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.VSKIN VSKIN1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8070
      Caption         =   "����-mid"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��������"
         Height          =   495
         Left            =   6360
         Style           =   1  '�׷���
         TabIndex        =   8
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "������ �� ������Ʈ�� ������ �߰��˴ϴ�."
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   1080
         TabIndex        =   7
         Top             =   3840
         Width           =   7695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "Will you take the train or fly? ������ Ÿ�� ���Ŵ� ������ ���Ŵ�?"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   720
         TabIndex        =   6
         Top             =   3240
         Width           =   8895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "Many people worry fly about war ���� ������� ���￡ ���� �����Ѵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   8775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "The first class begins at 9:00 1���ô� 9�ÿ� ���۵ȴ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1440
         TabIndex        =   4
         Top             =   2280
         Width           =   9015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "I have never seen such a fast bird �� �׷��� ���� ���� ���� �� ���� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   9015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "She loos just like her mother �׳�� ��Ӵϸ� �� ��Ҵ�."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   7335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "I always go to the same store �� �� ���� ���Կ� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   7095
      End
   End
End
Attribute VB_Name = "midst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell App.Path & "\" & "main.exe"
End
End Sub

