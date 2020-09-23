VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   " About Dynamic Menu Example"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   3960
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1110
      Left            =   0
      Picture         =   "frmAbout.frx":1042
      ScaleHeight     =   1050
      ScaleWidth      =   8055
      TabIndex        =   6
      Top             =   0
      Width           =   8115
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "skrodal@altavista.net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   4560
         MouseIcon       =   "frmAbout.frx":CFEC
         MousePointer    =   99  'Custom
         TabIndex        =   10
         ToolTipText     =   "Click to open default e-mail application"
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "members.dingoblue.net.au/~skrodal/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   4560
         MouseIcon       =   "frmAbout.frx":D13E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "Sorcery Creations' web-site"
         Top             =   840
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-m@il:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WWW:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   1080
      Width           =   8115
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   7020
         TabIndex        =   11
         Top             =   2280
         Width           =   855
      End
      Begin VB.Image i3 
         Height          =   675
         Left            =   7200
         MouseIcon       =   "frmAbout.frx":D290
         Picture         =   "frmAbout.frx":D3E2
         Top             =   480
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Image i2 
         Height          =   675
         Left            =   7200
         MouseIcon       =   "frmAbout.frx":DE45
         Picture         =   "frmAbout.frx":DF97
         Top             =   480
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thank you for downloading this code!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Hopefully you liked the code, and learned from it!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":EA0D
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":EAE8
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Regards from Sorcery Creations - Simon the Sorcerer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   4335
      End
      Begin VB.Image i1 
         Height          =   675
         Left            =   7200
         MouseIcon       =   "frmAbout.frx":EB8B
         Picture         =   "frmAbout.frx":ECDD
         Top             =   480
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Menu Array Example By Simon the Sorcerer - Sorcery Creations   '
'                                                                '
' One essential point which you must remember is that the        '
' mnuFilesInDir is an ARRAY. This does not happen automatically! '
' You need to set the index property of the menu item to 0.      '
' This is of course done for you in this example, nevertheless - '
' if you're a newbie - have a look at the menu editor and find   '
' out what I mean.                                               '
'                                                                '
' In this example, we learn how to take all the files in a set   '
' directory, strip the extension {.***) and then add the file    '
' names to a menu. I originally made this in order to add all    '
' favourite files (.url) in a favourites menu in a web-browser.  '
' Although you might not see the immediate need for this now,    '
' you will soon learn that it is essential to know the basics on '
' how to create a menu array, to load/unload controls etc. The   '
' string manipulation in the fillMenu() might also come in handy!'
'                                                                '
' Use this whatever and however you like - royalty free of course'
' -> it's just code.                                             '
'                                                                '
' I am happy to answer any questions you might have :)           '
'                                                                '
' E-mail: skrodal@altavista.net                                  '
' URL:    members.dingoblue.net.au/~skrodal/                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Opening of files & URLS
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Opening of files & URLS

Dim counter As Integer

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
counter = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End
End Sub

Private Sub Label10_Click()
ShellExecute hWnd, "open", "http://members.dingoblue.net.au/~skrodal/", vbNullString, vbNullString, 0 ' Open URL
End Sub

Private Sub Label3_Click()
Shell "Start.exe " & "mailto:skrodal@altavista.net?Subject=""Dynamic Menu Feedback""", 0  ' Open default email app
End Sub

Private Sub Timer1_Timer() ' Animate the Sorcery Creations Logo...

If counter = 0 Then
    i1.Visible = True
    i2.Visible = False
    i3.Visible = False
counter = counter + 1
ElseIf counter = 1 Then
    i1.Visible = False
    i2.Visible = True
    i3.Visible = False
counter = counter + 1
ElseIf counter = 2 Then
    i1.Visible = True
    i2.Visible = False
    i3.Visible = False
counter = counter + 1
ElseIf counter = 3 Then
    i1.Visible = False
    i2.Visible = False
    i3.Visible = True
counter = 0
End If
End Sub

