VERSION 5.00
Begin VB.Form load 
   BorderStyle     =   0  '����
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
   StartUpPosition =   3  'Windows �⺻��
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
MsgBox "���ʽ������� �۲��� �����մϴ�. ���α׷��� �߸� ������ Ǯ������ �۲� ��ġâ�� �߸� ��ġ���ּ���.", vbInformation, "�ȳ�"
Shell App.Path & "\" & "nanumsongeulssibut.exe"
SaveSetting "ch", "2", "3", "Text1.Text"
MsgBox "���α׷��� ����� �Ͽ��ֽʽÿ�", vbInformation, "�˸�"

End If
End Sub
