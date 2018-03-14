VERSION 5.00
Begin VB.Form setting 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.VSKIN VSKIN1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4471
      Caption         =   "Setting"
      Resize          =   0   'False
      Begin VB.CheckBox Check3 
         Caption         =   "¸ÞÀÎÁøÀÔ½Ã À½¾Ç²ô±â"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "ºñ¹Ð¹øÈ£ ÀúÀå"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¾ÆÀÌµðÀúÀå"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
   End
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
SaveSetting "id", "2", "3", Login.Text1.Text
If Check1.value = 0 Then
DeleteSetting "id", "2", "3"
Else
SaveSetting "id", "2", "3", Login.Text1.Text
End If
End Sub

Private Sub Check2_Click()
SaveSetting "pw", "2", "3", Login.Text2.Text
If Check2.value = 0 Then
DeleteSetting "pw", "2", "3"
Else
SaveSetting "pw", "2", "3", Login.Text2.Text
End If
End Sub


Private Sub Check3_Click()
SaveSetting "mu", "2", "3", Login.Text2.Text
If Check3.value = 0 Then
DeleteSetting "mu", "2", "3"
Else
SaveSetting "mu", "2", "3", Login.Text2.Text
End If
End Sub

Private Sub Form_Load()

If Not GetSetting("id", "2", "3") = "" Then
Check1.value = 1
Else
Check1.value = 0
End If

If Not GetSetting("pw", "2", "3") = "" Then
Check2.value = 1
Else
Check2.value = 0
End If

If Not GetSetting("mu", "2", "3") = "" Then
Check3.value = 1
Else
Check3.value = 0
End If
End Sub
