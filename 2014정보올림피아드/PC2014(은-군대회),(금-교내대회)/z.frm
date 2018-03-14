VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form ftpupload 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _extentx        =   11033
      _extenty        =   10186
      caption         =   "FTPUpload - AdminMod - PC2014"
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   720
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   240
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FTP Upload"
         Height          =   615
         Left            =   3840
         TabIndex        =   15
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   4320
         Width           =   4455
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Text            =   "html/users/"
         Top             =   3840
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   9
         Text            =   "159753fksp"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Text            =   "21"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Text            =   "dltkddlr789"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Text            =   "112.175.184.51"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.FileListBox File1 
         Height          =   1170
         Left            =   360
         System          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "FTP,PHP 관리자 이외 업로드(조작)금지"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   5160
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "LocalFile:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   4440
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   6000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   6120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "AppPath(LocalFile)"
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
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "RemoteFile:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "UserPW:"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "FTPPort:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "UserID:"
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   3360
         TabIndex        =   5
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "FTP IP:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   3000
         Width           =   735
      End
   End
End
Attribute VB_Name = "ftpupload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
