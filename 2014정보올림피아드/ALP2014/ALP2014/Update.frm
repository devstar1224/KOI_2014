VERSION 5.00
Begin VB.Form Update 
   BorderStyle     =   0  '����
   Caption         =   "Update"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9975
      Caption         =   "//Update// - ALP2014"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "������Ʈ"
         Enabled         =   0   'False
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
         Left            =   1680
         TabIndex        =   5
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "?.?.?"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1560
         TabIndex        =   4
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "�ֽŹ���"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�OL"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "1.0.0"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�"
            Size            =   48
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "a�ǿ��ǳ׸�OL"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WinHttp As New WinHttpRequest
Private update_new As Byte
Private update_now As Byte

Private Sub Command1_Click()
WinHttp.Open "GET", "http://1.247.97.112/ALP/ver.htm"
WinHttp.Send
update_new = StrConv(WinHttp.ResponseBody, vbUnicode)
update_now = "1.0.1"
If update_new = update_now Then
ElseIf update_new <> update_now Then
    If MsgBox("�ֽ� ������ �ٿ�������?" & vbCr & "��������: " & update_now & " �ֽŹ���: " & update_new & "", 32 + 4, "����") = vbYes Then
    Shell "explorer http://1.247.97.112/ALP/ALP2014.exe"
    Unload Me
    Else
    End If
Else
    MsgBox "�߸��� �����Դϴ�." & vbCr & "��������: " & update_now & " �ֽŹ���: " & update_new & "", 0 + 16, "����"
End If
End Sub

Private Sub Form_Load()
WinHttp.Open "GET", "http://1.247.97.112/ALP/ver.htm"
WinHttp.Send
Label4.Caption = StrConv(WinHttp.ResponseBody, vbUnicode)
If Label2.Caption = Label4.Caption Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub


