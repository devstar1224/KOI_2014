VERSION 5.00
Begin VB.Form consol 
   BorderStyle     =   0  '绝澜
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '拳搁 啊款单
   Begin Project1.VSKIN VSKIN1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      Caption         =   "Consol - AdminMod - PC2014 "
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         Caption         =   "包府磊 FTP"
         Height          =   615
         Left            =   4320
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "包府磊 林家芒"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "consol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
dothosting.Show
End Sub

Private Sub Command2_Click()
ftpupload.Show
End Sub
