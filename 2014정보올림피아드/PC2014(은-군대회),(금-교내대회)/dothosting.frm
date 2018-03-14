VERSION 5.00
Begin VB.Form dothosting 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      Caption         =   "Admin - PC2014"
      Resize          =   0   'False
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "159753fksp"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "dltkddlr789"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "dltkddlr789"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "159753fksp"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "dltkddlr789"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "http://dltkddlr789.dothome.co.kr"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "ftp://112.175.184.51"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "DB PW"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "DB ID"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "DB Name"
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "DB informaiton : MySQL5.1"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "Wepserver information : Apache2.2"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "PHP ver: PHP5.2"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "FTP PW"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "FTP ID"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "M.D address"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "WepServer IP"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Caption         =   "DB Name"
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "dothosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
