VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form result 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.VSKIN VSKIN1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12091
      Caption         =   "결과기록"
      Resize          =   0   'False
      Begin MSComctlLib.ListView L 
         Height          =   5775
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "맞은 횟수"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "틀린 횟수"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "통과 횟수"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "판정"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "날짜"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
LoadTxt
End Sub
Private Sub LoadTxt()
On Error Resume Next
Dim GetStr As String, SplitStr() As String, FF As Byte, i As Integer
FF = FreeFile

If DirResultFile = True Then
    Open App.Path & "\Result\Result.ini" For Binary Access Read As #FF
    GetStr = Space(LOF(FF))
    Get FF, , GetStr
    Close #FF
    
    SplitStr = Split(GetStr, vbCrLf)
    
    For i = 0 To UBound(SplitStr)
        L.ListItems.Add , , Replace(Split(SplitStr(i), "|")(0), vbLf, "")
        L.ListItems(i + 1).SubItems(1) = Split(SplitStr(i), "|")(1)
        L.ListItems(i + 1).SubItems(2) = Split(SplitStr(i), "|")(2)
        L.ListItems(i + 1).SubItems(3) = Split(SplitStr(i), "|")(3)
        L.ListItems(i + 1).SubItems(4) = Split(SplitStr(i), "|")(4)
    Next i
End If
End Sub

