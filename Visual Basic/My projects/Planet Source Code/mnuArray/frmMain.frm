VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Dynamic Menu Example"
   ClientHeight    =   2445
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "Empty &Menu"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Me&nu"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   3720
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "An example on how to use a dynamic menu, using an array. See code for more details!"
      Height          =   495
      Left            =   80
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuFilesInDir 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About..."
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
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


Private Sub cmdEmpty_Click()
emptyMenu 'Clear all menu items
End Sub

Private Sub cmdFill_Click()
fillMenu ' Add all items in file list
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
File1.Path = App.Path 'Set the file lists path to this example's dir
End Sub

Sub fillMenu()          ' This sub adds all items in the filelist to the menu and removes the file extension
Dim tempName As String  ' Temporarely stores each listitem in the loop below
emptyMenu               ' Make sure the menu is empty when we start

For i = 1 To File1.ListCount ' Loop from menuItem #1 to the number of items in file list

Load mnuFilesInDir(i)        ' Load a new menuitem

tempName = File1.List(i - 1) ' Set tempname equal to the current file name in the list

   mnuFilesInDir(i).Caption = Left(tempName, Len(tempName) - 4)  ' Set the caption of this menu item equal to
                                                                 ' the tempName, but remove the .*** extension
                                                                 ' by starting from the Left in tempName,
                                                                 ' counting the length (Len) of the string and
                                                                 ' then remove the 4 last characters (-4).

Next i                       ' Resume to next item


mnuFilesInDir(0).Visible = False ' Set the divider in menu to invisible.
                                 ' This menuitem was created at design time
                                 ' and can not be loaded/unloaded during run
                                 ' time. It is however needed to initialize
                                 ' the array of menu items, and that's why
                                 ' I added it at design time.
End Sub


Sub emptyMenu()                      ' This sub clears the menu completely, leaving only
                                     ' one item - you guessed right! The one created at
                                     ' design time. Remember - a menu item created at design
                                     ' time can not be unloaded/loaded -> only arrays from this
                                     ' item can be.
                                     
File1.Refresh                        ' Refresh the file list (in case files have been added/removed
                                     ' from the dir since the app started.
                                     
mnuFilesInDir(0).Visible = True      ' Make the 'parent' menu item visible. If this is done after the
                                     ' loop, you will find out that the last menu item will not be
                                     ' removed. Why?  Simply because there must be (at least) one visible
                                     ' menu item, and if the design-made(0) item is invisible, one of the
                                     ' run time items have to stay visible.
                                     
For i = 1 To mnuFilesInDir.Count - 1 ' A loop that removes (unloads) all run-time-created items in the menu
Unload mnuFilesInDir(i)              ' As you can see the loop starts on 1 (i = 1) --? item 0 is the divider
                                     ' that's made in design mode. If we tried to unload it, VB would complain.

Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAbout.Show 1, Me
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1, Me  ' Show about box...
End Sub
