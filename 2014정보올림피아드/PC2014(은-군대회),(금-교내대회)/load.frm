VERSION 5.00
Begin VB.Form Loading 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   4050
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "load.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7080
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   360
         Top             =   120
      End
      Begin VB.Label Label2 
         Caption         =   "Loading....."
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Powerful Clean"
         BeginProperty Font 
            Name            =   "����"
            Size            =   27.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2760
         TabIndex        =   4
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "load.frx":0802
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "���� 1.0.1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         TabIndex        =   2
         Top             =   3480
         Width           =   1185
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "�÷���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         TabIndex        =   3
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '������ ����
         Caption         =   "�� ��ǰ�� ���� ����ڿ��� ����� �㰡�Ǿ����ϴ�."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
If Not GetSetting("7", "8", "9") = "" Then
Main.Show
Unload Me
Else
agree.Show
Unload Me
End If
End Sub
