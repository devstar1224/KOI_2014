VERSION 5.00
Begin VB.Form agree 
   BorderStyle     =   0  '없음
   Caption         =   "Agree"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   Icon            =   "agree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5610
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      Caption         =   "agree - PC2014"
      Resize          =   0   'False
      Begin VB.CheckBox Check1 
         BackColor       =   &H00404040&
         Caption         =   "Agree"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "동의"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   $"agree.frx":0802
         ForeColor       =   &H8000000B&
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   5175
      End
   End
End
Attribute VB_Name = "agree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.value = 1 Then
SaveSetting "7", "8", "9", Label1.Caption
Main.Show
Unload Me
Else

End If
End Sub
