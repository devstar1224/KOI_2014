VERSION 5.00
Begin VB.Form gameword 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin Project1.VSKIN VSKIN1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13785
      Caption         =   "Game"
      Resize          =   0   'False
      Begin VB.CommandButton Command2 
         Caption         =   "¿Ï·á"
         Height          =   375
         Left            =   7920
         TabIndex        =   27
         Top             =   6120
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   1024
         Left            =   720
         Top             =   4680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "È®ÀÎ"
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Åõ¸í
         Caption         =   " ÃÊ ºÐ ½Ã"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "½Ã°£:"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   6720
         Width           =   735
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "¶æ:"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   21.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3360
         TabIndex        =   24
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "custom"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   20
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "however"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   19
         Left            =   7560
         TabIndex        =   22
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "wait"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   18
         Left            =   480
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "miss"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   17
         Left            =   8160
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "turn"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Index           =   16
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "need"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   15
         Left            =   5040
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "hard"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   14
         Left            =   8160
         TabIndex        =   17
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "example"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Index           =   13
         Left            =   4920
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "problem"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   12
         Left            =   1920
         TabIndex        =   15
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "nature"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   11
         Left            =   7080
         TabIndex        =   14
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "telephone"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Index           =   10
         Left            =   5640
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "taxi"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   9
         Left            =   7080
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "button"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Index           =   8
         Left            =   6600
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "all"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Index           =   7
         Left            =   3720
         TabIndex        =   10
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "bear"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   6
         Left            =   5880
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "after"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   5
         Left            =   3480
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "airport"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "bird"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   3
         Left            =   3960
         TabIndex        =   6
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "across"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   2
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "about"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   1
         Left            =   4800
         TabIndex        =   4
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "³ª´®¼Õ±Û¾¾ º×"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   375
      End
   End
End
Attribute VB_Name = "gameword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_MEMORY As Long = &H4
Private Const SND_NOSTOP As Long = &H10
Private Const SND_NOWAIT As Long = &H2000
Private Const SND_RESOURCE As Long = &H40004
Dim Check1 As Long
Dim Check2 As Long
Dim sec As Long
Dim mmm As Long
Dim hhh As Long
Dim result As String
Dim ssec As String
Dim smmm As String
Dim shhh As String


Private Sub Command1_Click()
If Text1.Text = "a" Then
Label22.Caption = "¶æ:ÇÏ³ªÀÇ"
Label1(0).Enabled = False
Else
End If

If Text1.Text = "about" Then
Label22.Caption = "¶æ:~°üÇÏ¿©"
Label1(1).Enabled = False
End If

If Text1.Text = "across" Then
Label22.Caption = "¶æ:~°¡·ÎÁú·¯"
Label1(2).Enabled = False
End If

If Text1.Text = "bird" Then
Label22.Caption = "¶æ:»õ"
Label1(3).Enabled = False
End If

If Text1.Text = "airport" Then
Label22.Caption = "¶æ:°øÇ×"
Label1(4).Enabled = False
End If

If Text1.Text = "after" Then
Label22.Caption = "¶æ:~ÈÄ¿¡"
Label1(5).Enabled = False
End If

If Text1.Text = "bear" Then
Label22.Caption = "¶æ:°õ"
Label1(6).Enabled = False
End If

If Text1.Text = "all" Then
Label22.Caption = "¶æ:¸ðµÎ"
Label1(7).Enabled = False
End If

If Text1.Text = "button" Then
Label22.Caption = "¶æ:¹öÆ°"
Label1(8).Enabled = False
End If

If Text1.Text = "taxi" Then
Label22.Caption = "¶æ:ÅÃ½Ã"
Label1(9).Enabled = False
End If

If Text1.Text = "telephone" Then
Label22.Caption = "¶æ:ÀüÈ­±â"
Label1(10).Enabled = False
End If

If Text1.Text = "nature" Then
Label22.Caption = "¶æ:ÀüÈ­±â"
Label1(11).Enabled = False
End If

If Text1.Text = "problem" Then
Label22.Caption = "¶æ:¹®Á¦"
Label1(12).Enabled = False
End If

If Text1.Text = "example" Then
Label22.Caption = "¶æ:¿¹"
Label1(13).Enabled = False
End If

If Text1.Text = "hard" Then
Label22.Caption = "¶æ:¾î·Á¿î"
Label1(14).Enabled = False
End If
 
If Text1.Text = "need" Then
Label22.Caption = "¶æ:ÇÊ¿äÇÑ"
Label1(15).Enabled = False
End If

If Text1.Text = "turn" Then
Label22.Caption = "¶æ:µ¹´Ù"
Label1(16).Enabled = False
End If

If Text1.Text = "miss" Then
Label22.Caption = "¶æ:~¾ç"
Label1(17).Enabled = False
End If

If Text1.Text = "wait" Then
Label22.Caption = "¶æ:±â´Ù¸®´Ù"
Label1(18).Enabled = False
End If

If Text1.Text = "however" Then
Label22.Caption = "¶æ:±×·¯³ª"
Label1(19).Enabled = False
End If

If Text1.Text = "custom" Then
Label22.Caption = "¶æ:°ü½À"
Label1(20).Enabled = False
End If

End Sub

Private Sub Command2_Click()
PlaySound App.Path & "\" & "sound\click.wav", 0, SND_ASYNC
Timer1.Enabled = False
sec = 0
mmm = 0
hhh = 0
ssec = sec
smmm = mmm
shhh = hhh
result = ssec + "ÃÊ" + smmm + "ºÐ" + shhh + "½Ã°£"
MsgBox Label24.Caption & "°¡ ÃÊ°ú µÇ¾ú½À´Ï´Ù. ¾÷µ¥ÀÌÆ®¸¶´Ù °ÔÀÓÀÌ ´Ù¾çÇØÁý´Ï´Ù.  "
Main.Show
Unload Me
End Sub


Private Sub Form_Load()
PlaySound App.Path & "\" & "sound\game\On the Bach.wav", 0, SND_ASYNC
Check2 = Format(Time, "ss")
Timer1.Enabled = True
End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label20_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
Check1 = Format(Time, "ss")

If Check1 = Check2 + 1 Then
sec = sec + 1

If sec = 60 Then
mmm = mmm + 1
sec = 0


If mmm = 60 Then
hhh = hhh + 1
mmm = 0
End If
End If
End If

Check2 = Check1
ssec = sec
smmm = mmm
shhh = hhh
result = ssec + "ÃÊ" + smmm + "ºÐ" + shhh + "½Ã°£"
Label24.Caption = result
End Sub

