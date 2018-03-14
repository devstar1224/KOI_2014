VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   0  '없음
   Caption         =   "Main"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   8895
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   13361
      Caption         =   "Powerful Clean 2014 - PC2014"
      Resize          =   0   'False
      Begin VB.CheckBox Check1 
         BackColor       =   &H00404040&
         Caption         =   "인터넷 속도 최적화"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00404040&
         Caption         =   "프로세스 메모리 최적화"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00404040&
         Caption         =   "키보드 반응속도 최적화"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00404040&
         Caption         =   "메뉴창 팝업속도 최적화"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404040&
         Caption         =   "시작"
         Height          =   495
         Left            =   600
         MaskColor       =   &H00808080&
         TabIndex        =   13
         Top             =   3480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Standard Mod"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   3720
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Game Mod"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   6960
         TabIndex        =   11
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "적용"
         Height          =   495
         Left            =   5400
         TabIndex        =   10
         Top             =   4320
         Width           =   2895
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00404040&
         Caption         =   "Windows 시작시 실행"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00404040&
         Caption         =   "프로그램 시작시 Standard Mod 적용"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00404040&
         Caption         =   "프로그램 시작시 Game Mod 적용"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "프로세스 최적화"
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "키보드 감도설정"
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "UpDate"
         Height          =   495
         Left            =   4920
         TabIndex        =   4
         Top             =   5880
         Width           =   3735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "컴퓨터 사용시간"
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "예약종료"
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   " 버그/문의 매일 보내기"
         Height          =   495
         Left            =   4920
         TabIndex        =   1
         Top             =   6600
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "종합 최적화"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "[대기중]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "[대기중]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "[대기중]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "[대기중]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3120
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "[시작 대기중]"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   4800
         X2              =   4800
         Y1              =   720
         Y2              =   7440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   4680
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "Mod Setting"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   26.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   5280
         TabIndex        =   22
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   240
         Y1              =   720
         Y2              =   7440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "Setting"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000A&
         X1              =   8760
         X2              =   8760
         Y1              =   720
         Y2              =   7440
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000A&
         X1              =   4920
         X2              =   8640
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "Tools"
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
         Left            =   2040
         TabIndex        =   20
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000A&
         DrawMode        =   12  'Nop
         X1              =   4920
         X2              =   8640
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "Program"
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
         Height          =   495
         Left            =   6120
         TabIndex        =   19
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000A&
         X1              =   360
         X2              =   4680
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "PHPVer : 5.2   Wepserver information:Apache 2.2"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   6840
         Width           =   4335
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check5_Click()
On Error Resume Next
If Check5.value = 1 Then
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
SaveSetting "7", "8", "8", Label1.Caption
Else
Set Reg = CreateObject("wscript.shell")
Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
DeleteSetting "7", "8", "8"
Check5.value = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If isUnloadfromDLL = True Then
        pbcon_Unload
    Else
        If MsgBox("프로그램을 종료하시겠습니까?", vbYesNo + vbQuestion) = vbYes Then
            pbcon_Unload
        Else
            Cancel = True
        End If
    End If
    
End Sub
Private Sub Check6_Click()


On Error Resume Next
If Check6.value = 1 Then
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\InternetSettings", "10"
moduleMemoryClean.Main
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardSpeed", "99"
Reg.RegWrite "HKEY_CURRENT_USER\ControlPanel\Desktop\MenuShowDelay", "0"
Option1.value = True
SaveSetting "7", "8", "7", Label1.Caption
Else
DeleteSetting "7", "8", "7"
Check6.value = 0
End If
End Sub

Private Sub Check7_Click()


On Error Resume Next
If Check7.value = 1 Then
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\InternetSettings", "10"
moduleMemoryClean.Main
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardSpeed", "99"
Reg.RegWrite "HKEY_CURRENT_USER\ControlPanel\Desktop\MenuShowDelay", "0"
Process.Command2.value = True
Option2.value = True
SaveSetting "7", "8", "6", Label1.Caption
Else
DeleteSetting "7", "8", "6"
Check7.value = 0
End If

End Sub

Private Sub Command1_Click()


On Error Resume Next
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")

If Check1.value = 1 Then
Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\InternetSettings", "10"
Label2.Caption = "[완료]"
Else
Label2.Caption = "[대기중]"
End If

If Check2.value = 1 Then
moduleMemoryClean.Main
Label3.Caption = "[완료]"
Else
Label3.Caption = "[대기중]"
End If

If Check3.value = 1 Then
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardSpeed", "99"
Label4.Caption = "[완료]"
Else
Label4.Caption = "[대기중]"
End If

If Check4.value = 1 Then
Reg.RegWrite "HKEY_CURRENT_USER\ControlPanel\Desktop\MenuShowDelay", "0"
Label5.Caption = "[완료]"
Else
Label5.Caption = "[대기중]"
End If
Label6.Caption = "[완료]"
End Sub

Private Sub Command10_Click()
support.Show
End Sub

Private Sub Command2_Click()


On Error Resume Next
If Option1.value = True Then
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\InternetSettings", "10"
moduleMemoryClean.Main
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardSpeed", "99"
Reg.RegWrite "HKEY_CURRENT_USER\ControlPanel\Desktop\MenuShowDelay", "0"
End If
If Option2.value = True Then
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\InternetSettings", "10"
moduleMemoryClean.Main
Reg.RegWrite "HKEY_CURRENT_USER\Control Panel\Keyboard\KeyboardSpeed", "99"
Reg.RegWrite "HKEY_CURRENT_USER\ControlPanel\Desktop\MenuShowDelay", "0"
Process.Command2.value = True
End If
End Sub

Private Sub Command3_Click()
MsgBox "하드에 불량섹터가 있는상테에서 프로세스 최적화를 할경우 F4 블루스크린이 뜹니다", , "주위"
Process.Show
End Sub

Private Sub Command4_Click()
key.Show
End Sub

Private Sub Command5_Click()
Help.Show
End Sub

Private Sub Command6_Click()
update.Show
End Sub

Private Sub Command7_Click()
time.Show
End Sub

Private Sub Command8_Click()
shutdown.Show
End Sub

Private Sub Command9_Click()
admin.Show
End Sub

Private Sub Form_Load()


On Error Resume Next
If Not GetSetting("7", "8", "8") = "" Then
Check5.value = 1
Else
Check5.value = 0
End If
If Not GetSetting("7", "8", "7") = "" Then
Check6.value = 1
Else
Check6.value = 0
End If
If Not GetSetting("7", "8", "6") = "" Then
Check7.value = 1
Else
Check7.value = 0
End If

End Sub

