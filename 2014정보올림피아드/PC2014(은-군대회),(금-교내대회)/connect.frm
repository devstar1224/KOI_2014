VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form connect 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin Project1.VSKIN VSKIN1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9763
      Caption         =   "Connect - Admin - PC2014"
      Begin VB.ListBox List1 
         Height          =   2220
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   5880
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Winsock2.Close
Winsock2.LocalPort = 1642
Winsock2.Listen
End Sub

Private Sub Winsock2_Close()
On Error Resume Next
 Sock(Index).Close
 
 For i = 0 To SockList.ListCount - 1
 
    If Split(SockList.List(i), ":")(1) = Index Then

    SockList.RemoveItem i
    NickList.RemoveItem i
    End If
DoEvents
 Next i
   LstUp
    
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept (requestID)
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim A As Long
A = windata
Winsock2.GetData windata
List1.AddItem A
End Sub
