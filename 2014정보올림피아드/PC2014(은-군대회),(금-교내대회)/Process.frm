VERSION 5.00
Begin VB.Form Process 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      Caption         =   "Process Clean - PC2014"
      Resize          =   0   'False
      Begin VB.ListBox List1 
         BackColor       =   &H00404040&
         ForeColor       =   &H80000005&
         Height          =   2400
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   5175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "정리하기"
         Height          =   615
         Left            =   3240
         TabIndex        =   2
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "새로고침"
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   3240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GWL_EXSTYLE = (-20)
Private Const ws_ex_layered = &H80000
Private Const LWA_ALPHA = &H2&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const fade As Integer = 8
Dim alpha As Integer
Dim exitable As Boolean
Dim Protect_Process() As String
Private Sub Command2_Click()
Dim i As Long, j As Long, CC As Boolean, Temp() As String

If MsgBox("작업하던 내용이 종료될 수 있습니다." & vbCrLf & "괜찮으십니까?", vbInformation + vbYesNo, "[알림]") = vbYes Then
    For i = 0 To List1.ListCount - 1
        For j = 0 To 22
            If List1.List(i) = Protect_Process(j) Then
                CC = True
            End If
        Next j
        
        If CC = False Then
           KillProcess List1.List(i)
        End If
        CC = False
    Next i
    
    List1.Clear
    Temp = Split(CollectionProccess, "%")
    
    For i = 1 To UBound(Temp)
        List1.AddItem Temp(i)
    Next i
End If
End Sub
Private Sub Command1_Click()
Dim Temp() As String, i As Long
Temp = Split(CollectionProccess, "%")

For i = 1 To UBound(Temp)
    List1.AddItem Temp(i)
Next i

ReDim Protect_Process(22)
Protect_Process(0) = "NateOnMain.exe"
Protect_Process(1) = "iexplorer.exe"
Protect_Process(2) = "explorer.exe"
Protect_Process(3) = "svchost.exe"
Protect_Process(4) = "ctfmon.exe"
Protect_Process(5) = "alg.exe"
Protect_Process(6) = "winlogon.exe"
Protect_Process(7) = "csrss.exe"
Protect_Process(8) = "smss.exe"
Protect_Process(9) = "spoolsv.exe"
Protect_Process(10) = "services.exe"
Protect_Process(11) = App.EXEName & ".exe"
Protect_Process(12) = "VB6.EXE"
Protect_Process(13) = "Skype.exe"
Protect_Process(14) = "spoolsv.exe"
Protect_Process(15) = "csrss.exe"
Protect_Process(16) = "dwm.exe"
Protect_Process(17) = "explorer.exe"
Protect_Process(18) = "lsass.exe"
Protect_Process(19) = "smss.exe"
Protect_Process(20) = "vds.exe"
Protect_Process(21) = "ctfmon.exe"
Protect_Process(22) = "rundll32.exe"
End Sub
Private Sub Form_Load()
Dim Temp() As String, i As Long
Temp = Split(CollectionProccess, "%")


For i = 1 To UBound(Temp)
    List1.AddItem Temp(i)
Next i

ReDim Protect_Process(22)
Protect_Process(0) = "NateOnMain.exe"
Protect_Process(1) = "iexplorer.exe"
Protect_Process(2) = "explorer.exe"
Protect_Process(3) = "svchost.exe"
Protect_Process(4) = "ctfmon.exe"
Protect_Process(5) = "alg.exe"
Protect_Process(6) = "winlogon.exe"
Protect_Process(7) = "csrss.exe"
Protect_Process(8) = "smss.exe"
Protect_Process(9) = "spoolsv.exe"
Protect_Process(10) = "services.exe"
Protect_Process(11) = App.EXEName & ".exe"
Protect_Process(12) = "VB6.EXE"
Protect_Process(13) = "Skype.exe"
Protect_Process(14) = "spoolsv.exe"
Protect_Process(15) = "csrss.exe"
Protect_Process(16) = "dwm.exe"
Protect_Process(17) = "explorer.exe"
Protect_Process(18) = "lsass.exe"
Protect_Process(19) = "smss.exe"
Protect_Process(20) = "vds.exe"
Protect_Process(21) = "ctfmon.exe"
Protect_Process(22) = "rundll32.exe"
End Sub

