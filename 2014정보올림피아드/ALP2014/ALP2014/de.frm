VERSION 5.00
Begin VB.Form de 
   BorderStyle     =   0  '����
   Caption         =   "de"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3201
      Caption         =   "//Resetting// - ALP2014"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "�ʱ�ȭ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  '��� ����
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "��Ʈ:"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "���� ��й�ȣ:"
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
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "de"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If GetSetting("1", "3", "pw") = Text1.Text Then
DeleteSetting "1", "3", "pw"
DeleteSetting "14", "key", "3"
MsgBox "�ʱ�ȭ�Ͽ����ϴ�.", vbInformation, "�˸�"
Unload Me
main.Show
Else
MsgBox "���� ��й�ȣ�� Ʋ���ϴ�", vbInformation, "�˸�"
End If
End Sub
Private Sub Form_Load()
Label2.Caption = "��Ʈ:" & GetSetting("1", "2", "hint")
End Sub

Private Sub Form_Unload(Cancel As Integer)
setting.Show
End Sub
