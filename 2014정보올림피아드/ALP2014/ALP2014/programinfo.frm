VERSION 5.00
Begin VB.Form programinfo 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7011
      Caption         =   "//Program information//  -  ALP2014"
      Resize          =   0   'False
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CCL����"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3720
         TabIndex        =   13
         Top             =   2520
         Width           =   2775
         Begin VB.Image Image1 
            Height          =   855
            Left            =   120
            Picture         =   "programinfo.frx":0000
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2565
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Support"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2760
         TabIndex        =   8
         Top             =   600
         Width           =   3735
         Begin VB.Label Label9 
            BackStyle       =   0  '����
            Caption         =   "Name: Lee-Sang IK"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "dltkddlr789@naver.com"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   11
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '����
            Caption         =   "wnddkd1224@nate.com"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   10
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '����
            Caption         =   "E-Mail: wnddkd1224@gmail.com"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   3375
         Begin VB.Label Label5 
            BackStyle       =   0  '����
            Caption         =   "ServerOS:Winodws7 32Bit"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "ProgramVer:1.0.0"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���۲�"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2415
         Begin VB.Label Label4 
            BackStyle       =   0  '����
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   "a�ǿ��ǳ׸�"
            BeginProperty Font 
               Name            =   "a�ǿ��ǳ׸�"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '����
            Caption         =   "a�ǿ��ǳ׸�OL"
            BeginProperty Font 
               Name            =   "a�ǿ��ǳ׸�OL"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "programinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
FadeIN Me
End Sub
Private Sub Form_Resize()
ctlSkin1.Height = Height
ctlSkin1.Width = Width
End Sub
Private Sub Form_Load()
mWidth = 3735
mHeight = 2055
gHW = Me.hWnd
Hook
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unhook
FadeOUT Me
End Sub

Private Sub Image1_Click()
MsgBox "Lee-Sang ik�� ���� �ۼ��� ALP2014��(��) ũ������Ƽ�� Ŀ���� ������ǥ��-�񿵸�-������� 4.0 ���� ���̼����� ���� �̿��� �� �ֽ��ϴ�." & vbCrLf & " �� ���̼����� ���� �̿��� �̿������ ��� ���ؼ��� dltkddlr789@naver.com�� �����Ͻʽÿ�.", vbInformation, "CCL"
End Sub

Private Sub Label1_Click()

End Sub
