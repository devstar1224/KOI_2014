VERSION 5.00
Begin VB.Form load 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   Icon            =   "load.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleWidth      =   210
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
Text1.Text = GetSetting("ch", "71", "3")
If Not GetSetting("ch", "71", "3") = "" Then
Shell App.Path & "\" & "ALP2014.exe"
End
Else
MsgBox "���ʽ������� �۲��� �����մϴ�. ���α׷��� �߸� ������ Ǯ������ �۲� ��ġâ�� �߸� ��ġ���ּ���.(�۲ü�ġâ�� ��2���� ��µ˴ϴ�.)", vbInformation, "�ȳ�"
Shell App.Path & "\" & "apop.exe"
Shell App.Path & "\" & "apopol.exe"
SaveSetting "ch", "71", "3", "Text1.Text"
MsgBox "���α׷��� ����� �Ͽ��ֽʽÿ�", vbInformation, "�˸�"
End
End If
End Sub
