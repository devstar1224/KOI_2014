VERSION 5.00
Begin VB.Form shutdown 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3625
      Caption         =   "Shutdown - PC2014"
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         Caption         =   "취소"
         Height          =   615
         Left            =   3000
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "설정"
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Text            =   "60=1분"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "시간:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "shutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "shutdown -s -t " & Text1.Text
End Sub

Private Sub Command2_Click()
Shell "shutdown -a"
End Sub


