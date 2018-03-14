VERSION 5.00
Begin VB.Form load 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   Icon            =   "load.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
AlwaysTop load, True
Text1.Text = GetSetting("ch", "2", "3")
If Not GetSetting("ch", "2", "3") = "" Then
Shell App.Path & "\" & "ES2014.exe"
End
Else
MsgBox "최초실행으로 글꼴을 구성합니다. 프로그램이 뜨면 압축을 풀으신후 글꼴 설치창이 뜨면 설치해주세요.", vbInformation, "안내"
Shell App.Path & "\" & "nanumsongeulssibut.exe"
SaveSetting "ch", "2", "3", "Text1.Text"
MsgBox "프로그램을 재시작 하여주십시오", vbInformation, "알림"

End If
End Sub
