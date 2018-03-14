VERSION 5.00
Begin VB.Form addwordtest 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame3 
      Appearance      =   0  '평면
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   6360
      TabIndex        =   11
      Top             =   -120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Cmd_RC 
         Caption         =   "확인"
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "결과"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   27.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label5 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   1200
         TabIndex        =   13
         Top             =   1560
         Width           =   3975
      End
   End
   Begin Project1.VSKIN VSKIN1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9975
      Caption         =   "추가단어테스트"
      Resize          =   0   'False
      Begin VB.CommandButton Cmd_Stop 
         Caption         =   "중지하기"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_C 
         Caption         =   "힌트"
         Height          =   495
         Index           =   2
         Left            =   4080
         TabIndex        =   6
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_C 
         Caption         =   "패스"
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_C 
         Caption         =   "확인"
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   4440
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "메인으로"
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Cmd_Start 
            Caption         =   "시작하기"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         X1              =   240
         X2              =   6000
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "정답 :"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   27.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2040
         TabIndex        =   1
         Top             =   2040
         Width           =   3855
      End
   End
End
Attribute VB_Name = "addwordtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Words() As String, Means() As String
Dim Pass As Integer, Bads As Integer, Goods As Integer
Dim Stage As Integer
Private Sub LoadTxt()
Dim FF As Byte, Temp As String, Temps() As String, LoadTemp As String, i As Integer
FF = FreeFile

If DirTextFile = True Then
    Open App.Path & "\Save\English.ini" For Binary As #FF
    Temp = Space(LOF(1))
    Get FF, , Temp
    LoadTemp = Temp
    Close #FF
    
    LoadTemp = Mid(LoadTemp, 1, Len(LoadTemp))
    Temps = Split(LoadTemp, vbCrLf)
    
    ReDim Words(UBound(Temps) - 1)
    ReDim Means(UBound(Temps) - 1)
    
    For i = 0 To UBound(Temps) - 1
        Words(i) = Split(Temps(i), "|")(0)
        Means(i) = Split(Temps(i), "|")(1)
    Next i
End If

End Sub
Private Sub Cmd_C_Click(Index As Integer)
If Index = 0 Then
    If Text1 = Words(Stage) Then
        MsgBox "정답입니다!", vbInformation, "[알림]"
        Goods = Goods + 1
    Else
        MsgBox "틀렸습니다!", vbCritical, "[알림]"
        Bads = Bads + 1
    End If
    
    If UBound(Words) = Stage Then
        MsgBox "테스트가 끝이 났습니다!" & vbCrLf & "결과창으로 갑니다!", vbInformation, "[알림]"
        EnResult
        Exit Sub
    End If
    Stage = Stage + 1
    
    Label2 = Means(Stage)
    Label4 = Stage + 1 & " / " & UBound(Words) + 1
    Text1 = ""
    If Cmd_C(2).Enabled = False Then Cmd_C(2).Enabled = True
ElseIf Index = 1 Then
    Pass = Pass + 1
    
    If UBound(Words) = Stage Then
        MsgBox "테스트가 끝이 났습니다!" & vbCrLf & "결과창으로 갑니다!", vbInformation, "[알림]"
        EnResult
        Exit Sub
    End If
    Stage = Stage + 1
    
    Label2 = Means(Stage)
    Label4 = Stage + 1 & " / " & UBound(Words) + 1
    Text1 = ""
    If Cmd_C(2).Enabled = False Then Cmd_C(2).Enabled = True
ElseIf Index = 2 Then
    ' 힌트 메시지박스로 띄운다.
    Dim HintStr As String, HintGrid(0 To 1) As Integer
    
    HintStr = Words(Stage)
    Randomize
    HintGrid(0) = CLng(Len(HintStr) * Rnd)
    HintGrid(1) = CLng((Len(HintStr) - 1) * Rnd)
        If HintGrid(1) >= HintGrid(0) Then HintGrid(1) = HintGrid(1) + 1
    HintStr = String(HintGrid(0), "■") & Mid$(HintStr, HintGrid(0) + 1, 1) & String$(Len(Words(Stage)) - HintGrid(0) - 1, "■")
    'HintStr = Mid$(HintStr, 1, HintGrid(0)) & "*" & Mid$(HintStr, HintGrid(0) + 2, Len(HintStr) - HintGrid(0) + 1)
    HintStr = Mid$(HintStr, 1, HintGrid(1)) & Mid$(Words(Stage), HintGrid(1) + 1, 1) & Mid$(HintStr, HintGrid(1) + 2, Len(HintStr) - HintGrid(1) + 1)
    
    MsgBox "힌트: " & HintStr, vbInformation, "단어 힌트"
    Cmd_C(2).Enabled = False
End If
End Sub
Private Sub Cmd_RC_Click()
Frame3.Visible = False
Ends
Main.Show
Unload Me
End Sub
Private Sub Cmd_Start_Click()
Stage = 0
Start
Label2 = Means(Stage)
Frame1.Enabled = False
Label4 = Stage + 1 & " / " & UBound(Words) + 1
Text1 = ""
End Sub
Private Sub Cmd_Stop_Click()
Ends
End Sub

Private Sub Command1_Click()
Main.Show
Unload Me
End Sub

Private Sub Form_Load()
LoadTxt
Frame3.Left = 60
Frame3.Top = 30
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Cmd_C_Click(0)
End Sub
Private Sub Start()
Dim i As Integer

Text1.Enabled = True
For i = 0 To 2
    Cmd_C(i).Enabled = True
Next i
Label2 = ""
Label4 = ""
End Sub
Private Sub Ends()
Dim i As Integer

Text1.Enabled = False
For i = 0 To 2
    Cmd_C(i).Enabled = False
Next i
Label2 = ""
Label4 = ""
Frame1.Enabled = True

SaveTxtResult

Pass = 0
Bads = 0
Goods = 0

End Sub
Private Sub EnResult()
Frame3.Visible = True
Bads = Bads + Pass

Label5 = "맞힌 수 : " & Goods & " / " & UBound(Words) + 1 & vbCrLf & _
"틀린 수 : " & Bads & " / " & UBound(Words) + 1 & vbCrLf & _
"통과 수 : " & Pass & " / " & UBound(Words) + 1 & vbCrLf

If Goods = UBound(Words) + 1 Then
    Label5 = Label5 & "판정 : Exellent!"
    Exit Sub
ElseIf Goods = 0 Then
    Label5 = Label5 & "판정 : Terrible!"
    Exit Sub
ElseIf Goods > (UBound(Words) + 1) / 2 And Goods <> UBound(Words) + 1 Then
    Label5 = Label5 & "판정 : Good!"
    Exit Sub
ElseIf Goods < (UBound(Words) + 1) / 2 And Goods <> 0 Then
    Label5 = Label5 & "판정 : Bad!"
    Exit Sub
ElseIf Goods = (UBound(Words) + 1) / 2 Then
    Label5 = Label5 & "판정 : Not Bad!"
End If

Cmd_C(0).Enabled = False
End Sub
Private Sub SaveTxtResult()
On Error Resume Next
' -- 결과를 텍스트로 저장
Dim FF As Byte, Temp As String, Temps() As String, Tempa As String, i As Integer

FF = FreeFile

MkDir App.Path & "\Result"
If DirResultFile = False Then
        
    Temp = Goods & "/" & UBound(Words) + 1 & "|" & _
    Bads & "/" & UBound(Words) + 1 & "|" & _
    Pass & "/" & UBound(Words) + 1 & "|" & _
    Mid(Split(Split(Label5, vbCrLf)(3), ":")(1), 2, Len(Split(Split(Label5, vbCrLf)(3), ":")(1)) - 1) & "|" & _
    Date
    
    Open App.Path & "\Result\Result.ini" For Binary Access Write As #FF
        Put #FF, , Temp
    Close #FF
    Exit Sub
Else
    Open App.Path & "\Result\Result.ini" For Binary Access Read As #FF
        Temp = Space(LOF(FF))
        Get FF, , Temp
    Close #FF
    
    Temps = Split(Temp, vbCrLf)

    For i = 0 To UBound(Temps)
        Tempa = Tempa & Temps(i) & vbCrLf
    Next i
    
    Tempa = Mid(Tempa, 1, Len(Tempa) - 1)
    Tempa = Tempa & Goods & "/" & UBound(Words) + 1 & "|" & _
    Bads & "/" & UBound(Words) + 1 & "|" & _
    Pass & "/" & UBound(Words) + 1 & "|" & _
    Mid(Split(Split(Label5, vbCrLf)(3), ":")(1), 2, Len(Split(Split(Label5, vbCrLf)(3), ":")(1)) - 1) & "|" & _
    Date
    
    Open App.Path & "\Result\Result.ini" For Binary Access Write As #FF
        Put #FF, , Replace(Tempa, vbCr, vbCrLf)
    Close #FF
    
End If
End Sub



