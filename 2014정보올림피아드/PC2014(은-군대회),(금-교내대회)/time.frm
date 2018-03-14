VERSION 5.00
Begin VB.Form time 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      Caption         =   "Use time - PC2014"
      Resize          =   0   'False
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2160
         Top             =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Label2"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "합계시간:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long

Private Declare Function DebugActiveProcessStop Lib "kernel32" (ByVal dwProcessId As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
   ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
   ByVal wFlags As Long) As Long
   
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Private Sub Timer1_Timer()
Label2.Caption = GetTickCount \ 1000 & "초 입니다"
End Sub
