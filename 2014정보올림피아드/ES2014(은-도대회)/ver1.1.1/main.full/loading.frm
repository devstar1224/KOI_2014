VERSION 5.00
Begin VB.Form loading 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.VSKIN VSKIN1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   -360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7858
      Caption         =   "VSKIN"
      Resize          =   0   'False
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   1920
         Top             =   3000
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "�÷���"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   5520
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "�� ��ǰ�� ���� ����ڿ��� ����� �㰡�Ǿ����ϴ�."
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "Ver:1.1.1"
         BeginProperty Font 
            Name            =   "�����ձ۾� ��"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5400
         TabIndex        =   1
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   2130
         Left            =   120
         Picture         =   "loading.frx":0000
         Top             =   1200
         Width           =   8505
      End
      Begin VB.Image Image1 
         Height          =   1560
         Left            =   720
         Picture         =   "loading.frx":21E3
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1560
      End
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Login.Show
Unload Me
End Sub

