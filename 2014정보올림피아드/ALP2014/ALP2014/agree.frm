VERSION 5.00
Begin VB.Form agree 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   Icon            =   "agree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8535
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   4575
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8535
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   1320
         Top             =   3600
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   8520
         X2              =   8520
         Y1              =   0
         Y2              =   4560
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   8520
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   8520
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   6840
         Picture         =   "agree.frx":038A
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1245
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '����
         Caption         =   "Ver:1.0.0"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   13
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '����
         Caption         =   "(ALP2014)"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "Automatic Login Program"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   11
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Image Image3 
         Height          =   1680
         Left            =   240
         Picture         =   "agree.frx":C6AF
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2400
      End
   End
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8070
      Caption         =   "//Agree// Automatic Login Program - ALP2014"
      Resize          =   0   'False
      Begin VB.Label Label10 
         BackStyle       =   0  '����
         Caption         =   "�� �ڵ�����Ǵ°��� ����(���)�Ͻðٽ��ϱ�?"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "������Ʈ��(Registry)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         Caption         =   "�����α׷��� "
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "�� �ڵ�����˴ϴ�."
         Height          =   255
         Left            =   6000
         TabIndex        =   6
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "������Ʈ��(Registry)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         ToolTipText     =   $"agree.frx":CA39
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   $"agree.frx":CB6B
         Height          =   615
         Left            =   720
         TabIndex        =   4
         Top             =   1560
         Width           =   7095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "���α׷� (ALP2014) �̿��� ���� ����"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�OL"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "����(Dissent)"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   2
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "   ����(Agree)"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   1035
         Left            =   4920
         Picture         =   "agree.frx":CC39
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   1050
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   1800
         Picture         =   "agree.frx":D03B
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1170
      End
   End
End
Attribute VB_Name = "agree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
MsgBox "���� ������ ���� �ϼ˽��ϴ�.", vbInformation, "�ȳ�"
SaveSetting "agree", "2", "3", "aac"
main.Show
Unload Me
End Sub

Private Sub Image2_Click()
MsgBox "�����ϼ˽��ϴ�. ���α׷��� �ֱ���� ���Ͽ� ���α׷��� �����մϴ�", vbInformation, "�ȳ�"
End
End Sub

Private Sub Label6_Click()
MsgBox "����(registry)��� ������� ���� �ü�� �ȿ��� �۵��ϴ� ��� ���α׷��� �ý��� ������ ��� �ִ� �����ͺ��̽��̴�. �ش� �ý��ۿ� ���� ���μ����� ����, �ֱ����ġ�� �뷮, ���ӵ� �ֺ���ġ�� ����, �ý��� �Ű�����, �������Ʈ����� ����ϴ� ������ Ÿ�԰� ���� �Ű�����(parameter) ���� ���� �ִ�.", vbInformation, "����"
End Sub

Private Sub Label9_Click()
MsgBox "����(registry)��� ������� ���� �ü�� �ȿ��� �۵��ϴ� ��� ���α׷��� �ý��� ������ ��� �ִ� �����ͺ��̽��̴�. �ش� �ý��ۿ� ���� ���μ����� ����, �ֱ����ġ�� �뷮, ���ӵ� �ֺ���ġ�� ����, �ý��� �Ű�����, �������Ʈ����� ����ϴ� ������ Ÿ�԰� ���� �Ű�����(parameter) ���� ���� �ִ�.", vbInformation, "����"
End Sub

Private Sub List1_Click()
End Sub

Private Sub Timer1_Timer()
If GetSetting("agree", "2", "3") = "aac" Then
If GetSetting("14", "key", "3") = "qq" Then
login.Show
agree.Hide
Timer1.Enabled = False
Else
main.Show
agree.Hide
Timer1.Enabled = False
End If
End If
Frame1.Left = 99999
End Sub
