VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ftpupload 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.VSKIN VSKIN1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   10186
      Caption         =   "FTPUpload - AdminMod - PC2014"
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   2640
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3000
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FTP Upload"
         Height          =   615
         Left            =   3840
         TabIndex        =   15
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   4320
         Width           =   4455
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Text            =   "html/users/"
         Top             =   3840
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   9
         Text            =   "159753fksp"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Text            =   "21"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Text            =   "dltkddlr789"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Text            =   "112.175.184.51"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   360
         System          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FTP,PHP 관리자 이외 업로드(조작)금지"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   5160
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "LocalFile:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   4440
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   6000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   6120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "AppPath(LocalFile)"
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
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "RemoteFile:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "UserPW:"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "FTPPort:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "UserID:"
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   3360
         TabIndex        =   5
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FTP IP:"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   3000
         Width           =   735
      End
   End
End
Attribute VB_Name = "ftpupload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type Com
    Reply As String
    BackCommand As String
End Type

Dim Commun(5) As Com
Dim CommunState As Integer
Dim Site As String
Dim Port As String
Dim Username As String
Dim Password As String
Dim RemoteFile As String
Dim LocalFile As String
Dim Buffersize As Long
Dim CloseAfterSend As Boolean

Dim bTimeOut As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" ( _
                                          ByVal hFtpSession As Long, _
                                          ByVal lpszSearchFile As String, _
                                          lpFindFileData As WIN32_FIND_DATA, _
                                          ByVal dwFlags As Long, _
                                          ByVal dwContent As Long) As Long

Private Sub CheckTimeout(ByVal wTime As Single)   '!
    
    Dim oTime!, sTime!
    
    bTimeOut = False
    
    oTime! = Timer
    Do
        If Timer < oTime! Then
           oTime! = oTime! - 86400
        End If
        sTime! = Timer - oTime!
        Sleep 1
        DoEvents
        If bTimeOut = True Then
            Exit Sub
        End If
    Loop Until wTime! < sTime!
    
    MsgBox "ftp connect timeout!!"
    Close #1
    
End Sub
Private Sub Command1_Click()

    Dim Nr1 As Integer
    Dim Nr2 As Integer
    Dim LocalIP As String
    
    Site = Text1.Text
    Port = Text3.Text
    Username = Text2.Text
    Password = Text4.Text
    LocalFile = Text6.Text
    RemoteFile = Text5.Text
    
    Commun(0).Reply = "220"
    Commun(0).BackCommand = "USER " + Username
    
    Commun(1).Reply = "331"
    Commun(1).BackCommand = "PASS " + Password
    
    Commun(2).Reply = "230"
    Commun(2).BackCommand = "TYPE I"
    
    Commun(3).Reply = "200"
    Commun(3).BackCommand = "PORT"
    
    Commun(4).Reply = "200"
    Commun(4).BackCommand = "STOR " + RemoteFile
    
    Commun(5).Reply = ""
    Commun(5).BackCommand = ""
    
    Buffersize = 2920
    LocalIP = Winsock1.LocalIP
    
    Do Until InStr(LocalIP, ".") = 0
        LocalIP = Left(LocalIP, InStr(LocalIP, ".") - 1) + "," + Right(LocalIP, Len(LocalIP) - InStr(LocalIP, "."))
    Loop
        
    Randomize Timer
    Nr1 = Int(Rnd * 12) + 5
    Nr2 = Int(Rnd * 254) + 1
    Commun(3).BackCommand = "PORT " + LocalIP + "," + Trim(Str(Nr1)) + "," + Trim(Str(Nr2))
    
    Winsock2.Close
    Do Until Winsock2.State = 0
        DoEvents
    Loop
    
    Winsock2.LocalPort = (Nr1 * 256) + Nr2
    Winsock2.Listen
    
    Winsock1.Close
    Do Until Winsock1.State = 0
        DoEvents
    Loop
    
    CommunState = 0
    
    Winsock1.RemoteHost = Site
    Winsock1.RemotePort = Port
    
    Winsock1.Connect
    Do Until Winsock1.State = 7 Or Winsock1.State = 9
        DoEvents
    Loop
    
    Select Case Winsock1.State
        Case 9
            MsgBox "Couldn't reach server " + Site + ".", vbOKOnly + vbInformation, "FTP Upper"
        Case 7
            Open LocalFile For Binary As #1
    End Select
    
    Call CheckTimeout(5)

End Sub

Private Sub File1_Click()
Text6.Text = File1.Path & "\" & File1.FileName
Text5.Text = "html/users/" & File1.FileName
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim tmpS As String
        
    bTimeOut = True
    
    Winsock1.GetData tmpS, , bytesTotal
    Debug.Print tmpS;
    
    Select Case Left(tmpS, 3)
        
        Case Commun(CommunState).Reply
            Winsock1.SendData Commun(CommunState).BackCommand + Chr(13) + Chr(10)
            Debug.Print Commun(CommunState).BackCommand
            CommunState = CommunState + 1
        Case "150"
            Do Until Winsock2.State = 7
                DoEvents
            Loop
            SendNextData
        Case "226"
            Winsock1.Close
            Do Until Winsock1.State = 0
                DoEvents
            Loop
            MsgBox "Transfer complete.", vbOKOnly + vbInformation, "알림"
        Case Else
    
            MsgBox "Bad reply: " + Left(tmpS, Len(tmpS) - 2), vbOKOnly + vbInformation, "FTP Upper"
            Close #1
    End Select

End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)

Winsock2.Close
Do Until Winsock2.State = 0
    DoEvents
Loop

Winsock2.Accept requestID
Do Until Winsock2.State = 7
    DoEvents
Loop

End Sub

Private Sub SendNextData()
On Error Resume Next
    Dim Take As Long
    Dim Buffer() As Byte
    
    If LOF(1) - Seek(1) < Buffersize Then
        Take = LOF(1) - Seek(1) + 1
    Else
        Take = Buffersize
    End If
    
    ReDim Buffer(0 To Take - 1)
    
    Get #1, , Buffer
    Winsock2.SendData Buffer
    Erase Buffer
    
    Label1 = Trim(Str(Seek(1)) - 1) + "/" + Trim(Str(LOF(1)))
    
    If Take < Buffersize Then
        Close #1
        CloseAfterSend = True
    End If

End Sub

Private Sub Winsock2_SendComplete()
    
    If CloseAfterSend = True Then
        Winsock2.Close
            
        Do Until Winsock2.State = 0
            DoEvents
        Loop
        CloseAfterSend = False
    Else
        SendNextData
    End If

End Sub


