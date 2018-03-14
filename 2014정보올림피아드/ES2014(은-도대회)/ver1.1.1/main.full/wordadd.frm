VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form wordadd 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin Project1.VSKIN 단어추가 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11880
      Caption         =   "단어 추가"
      Resize          =   0   'False
      Begin VB.CommandButton Cmd_Add 
         Caption         =   "추가하기"
         Height          =   495
         Left            =   3840
         TabIndex        =   8
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton Cmd_Re 
         Caption         =   "삭제하기"
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton Cmd_Rep 
         Caption         =   "수정하기"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   6240
         Width           =   1695
      End
      Begin VB.TextBox txt2 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   5760
         Width           =   4095
      End
      Begin VB.TextBox txt1 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   5160
         Width           =   4095
      End
      Begin MSComctlLib.ListView L 
         Height          =   4455
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "영어"
            Object.Width           =   2681
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "단어"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "날짜"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "단어 뜻:"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "단어:"
         BeginProperty Font 
            Name            =   "나눔손글씨 붓"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   5160
         Width           =   1095
      End
   End
End
Attribute VB_Name = "wordadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sindex As Integer, Change As Boolean
Private Sub Cmd_Add_Click()
If Trim$(txt1) = "" Or Trim$(txt2) = "" Then
    MsgBox "정보를 기입해 주세요!"
    Exit Sub
End If

If Change = True Then
    L.ListItems.Remove Sindex
    L.ListItems.Add Sindex, , txt1
    L.ListItems(Sindex).SubItems(1) = txt2
    L.ListItems(Sindex).SubItems(2) = Date
    txt1 = ""
    txt2 = ""
    txt1.SetFocus
    Change = False
    Exit Sub
End If

L.ListItems.Add , , txt1
L.ListItems(L.ListItems.Count).SubItems(1) = txt2
L.ListItems(L.ListItems.Count).SubItems(2) = Date

txt1 = ""
txt2 = ""
txt1.SetFocus
End Sub
Private Sub Cmd_Re_Click()
On Error GoTo Err
L.ListItems.Remove L.SelectedItem.Index

Exit Sub
Err:
MsgBox "리스트 목록을 선택해 주세요!", vbCritical, "[알림]"
End Sub
Private Sub Cmd_Rep_Click()
On Error GoTo Err
Sindex = L.SelectedItem.Index

txt1 = L.ListItems(Sindex).Text
txt2 = L.ListItems(Sindex).SubItems(1)
Change = True

Exit Sub
Err:
MsgBox "리스트 목록을 선택해 주세요!", vbCritical, "[알림]"
End Sub
Private Sub Form_Load()
SaveAndLoadTxt "Load"
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveAndLoadTxt "Save"
End Sub
Private Sub SaveAndLoadTxt(Order As String)
On Error Resume Next
Dim FF As Byte, SaveTemp As String, i As Integer, LoadTemp As String, LoadTemps() As String
Dim Temp As String
FF = FreeFile

If Order = "Load" Then
    If DirTextFile = True Then
        Open App.Path & "\Save\English.ini" For Binary As #FF
        Temp = Space(LOF(1))
        Get FF, , Temp
        LoadTemp = Temp
        Close #FF
     
        LoadTemp = Mid(LoadTemp, 1, Len(LoadTemp))
        LoadTemps = Split(LoadTemp, vbCrLf)
        
        For i = 0 To UBound(LoadTemps) - 1
            L.ListItems.Add , , Split(LoadTemps(i), "|")(0)
            L.ListItems(L.ListItems.Count).SubItems(1) = Split(LoadTemps(i), "|")(1)
            L.ListItems(L.ListItems.Count).SubItems(2) = Split(LoadTemps(i), "|")(2)
        Next i
    End If
ElseIf Order = "Save" Then
    For i = 1 To L.ListItems.Count
        SaveTemp = SaveTemp & L.ListItems(i).Text & "|" & _
        L.ListItems(i).ListSubItems(1) & "|" & _
        L.ListItems(i).ListSubItems(2) & vbCrLf
    Next i
    SaveTemp = Mid(SaveTemp, 1, Len(SaveTemp) - 2)
    
    MkDir App.Path & "\Save"
    
    Open App.Path & "\Save\English.ini" For Output As #FF
        Print #FF, SaveTemp
    Close #FF
    
End If
End Sub

Private Sub Txt2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Add_Click
End If
End Sub


