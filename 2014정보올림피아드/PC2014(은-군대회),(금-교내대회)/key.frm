VERSION 5.00
Begin VB.Form key 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4048
      Caption         =   "Key"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "적용"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "감도"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   495
      End
   End
End
Attribute VB_Name = "key"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IsNumeric(Text1.Text) = False Then
MsgBox "숫자만 입력해주세요.", , "경고"
Else
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardSpeed", Text1.Text
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardDelay", "0"
MsgBox "성공적으로 적용하였습니다", , "알림"
End If
End Sub
