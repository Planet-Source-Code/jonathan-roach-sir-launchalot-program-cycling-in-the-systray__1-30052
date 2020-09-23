VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuite 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Timer tmrOnScreen 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   5520
   End
   Begin VB.PictureBox picHeaderIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3120
      Picture         =   "Lalot.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox picHeaderIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2520
      Picture         =   "Lalot.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox picHeaderIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "Lalot.frx":0614
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox picHeaderIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1320
      Picture         =   "Lalot.frx":091E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox picHeaderIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "Lalot.frx":0C28
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox picHeaderIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox picTrayIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   1920
      Picture         =   "Lalot.frx":0F32
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   240
   End
   Begin VB.PictureBox picTrayIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   1560
      Picture         =   "Lalot.frx":107C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   240
   End
   Begin VB.PictureBox picTrayIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   1200
      Picture         =   "Lalot.frx":11C6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   240
   End
   Begin VB.PictureBox picTrayIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   840
      Picture         =   "Lalot.frx":1310
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   240
   End
   Begin VB.PictureBox picTrayIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "Lalot.frx":145A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   240
   End
   Begin VB.PictureBox picTrayIcons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   240
   End
   Begin VB.PictureBox PicHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.PictureBox picHeaderIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblHeader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4710
         TabIndex        =   1
         Top             =   150
         Width           =   105
      End
   End
   Begin VB.Frame frameBy 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   3600
      Width           =   5055
      Begin VB.Label lblBy 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "By Jonathan Roach"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "Email: sdsupport@gto.net"
         Top             =   195
         Width           =   3210
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Lalot.frx":15A4
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Lalot.frx":1656
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   4815
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu itmMain 
         Caption         =   "&Sir Launchalot Main"
      End
      Begin VB.Menu itmExplorer 
         Caption         =   "&Windows Explorer"
      End
      Begin VB.Menu itmCpl 
         Caption         =   "&Control Panel"
      End
      Begin VB.Menu itmIE 
         Caption         =   "&Internet Explorer"
      End
      Begin VB.Menu itmEmail 
         Caption         =   "&Email"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code: Sir Launchalot
'Author: Jonathan Roach - sdsupport@gto.net
'Purpose: To demonstrate, cycled program launching from systray
'
'Details: Only tested on Win98 and Win2K so far, developed under VB5 Pro
'
'
'Comments/Votes are always appreciated
'
'API declaration for launching the various programs
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim CurrentSelection As Integer 'Will hold the currently selected program
Dim SelText As String           'The title for the current selection
Dim LaunchText As String        'The actual program name/path to launch
Dim TipText As String           'Tooltip text for various states of program

'Constants used to define the form size
Private Const Expanded As Long = 4185
Private Const Compact As Long = 870

Private Sub cmdQuite_Click()
'Kill the icon in the tray and unload, then we end.
Shell_NotifyIcon NIM_DELETE, icoDat
Unload Form1
End
End Sub

Private Sub Form_Load()
'Set the default height for the form and show it
Form1.Height = Compact
Me.Show
'Set the current selection to the main Sir Launchalot screen
CurrentSelection = 1
'Call the SelectionState sub to setup details based on the
'current selection
SelectionState
LoadIcons True
End Sub

Private Sub SelectionState()
'This sub determines the currently selected program/item
'and sets the tool tip text, the form height as well
'as the program to launch and it's caption for the form
tmrOnScreen.Enabled = False
Form1.Height = Compact
Select Case CurrentSelection
Case 1 'Default (Configuration Screen)
'I will only comment this case statement as the others
'are doing the same things.
'
    'Set the text to be displayed in the form
    SelText = "Sir Launchalot"
    'Set the height of the form to expanded view size
    Form1.Height = Expanded
    'Set the path to the program we want to launch
    'since this is the main screen of my demo, we do
    'not launch anything
    LaunchText = "Nothing"
    'Set the tooltip text for the icon in the systray
    'depending on which program/item is ready
    TipText = "Sir Launchalot Demo"
Case 2 'Windows Explorer
    SelText = "Windows Explorer"
    'Start our timer, which will display the form briefly
    'showing the loaded program icon and name, then disappear.
    tmrOnScreen.Enabled = True
    LaunchText = "explorer.exe"
    TipText = "Windows explorer ready - Left click to start!"
Case 3 'Control Panel
    SelText = "Control Panel"
    tmrOnScreen.Enabled = True
    LaunchText = "Control.exe"
    TipText = "Control panel ready - Left click to start!"
Case 4 'Internet Explorer
    SelText = "Internet Explorer"
    tmrOnScreen.Enabled = True
    LaunchText = "iexplore.exe"
    TipText = "Internet Explorer ready - Left click to start!"
Case 5 'Email
    SelText = "Email The Author"
    tmrOnScreen.Enabled = True
    LaunchText = "mailto:sdsupport@gto.net?subject=Sir Launchalot Code"
    TipText = "Email to the code author - Left click to start!"
End Select
'Set the caption of the forms header label to the value
'of SelText, determined in the above case structure
lblHeader.Caption = SelText
'Show the form
Me.Show
End Sub

Private Sub LoadIcons(FirstLoad As Boolean)
'First the Header Icon
picHeaderIcon.Picture = picHeaderIcons(CurrentSelection).Picture
'Then the tray
'Setup the data for our icon in the systray
icoDat.cbSize = Len(icoDat)
icoDat.hWnd = Form1.hWnd
icoDat.uId = vbNull
icoDat.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
icoDat.uCallBackMessage = WM_MOUSEMOVE
icoDat.hIcon = picTrayIcons(CurrentSelection).Picture
icoDat.szTip = TipText & vbNullChar

'Call the Shell_NotifyIcon function to add the icon to the taskbar
'status area, If FirstLoad is True then we are adding an icon
'for the first time, otherwise we are modifying the current icon.
If FirstLoad = True Then
icoDat.szTip = "Sir Launchalot Demo" & vbNullChar
Shell_NotifyIcon NIM_ADD, icoDat
Else
Shell_NotifyIcon NIM_MODIFY, icoDat
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The following code intercepts the mouse activity on the
'systray icon and acts accordingly depending on the buttons.
Dim msg As Long
msg = X / Screen.TwipsPerPixelX
Select Case msg
    'If it is the left button then we launch the program currently active
    Case WM_LBUTTONDOWN
        ShellExecute Me.hWnd, "Open", LaunchText, vbNullString, vbNullString, vbNormalFocus
    Case WM_LBUTTONUP
    Case WM_LBUTTONDBLCLK
    'If it is the right button we cycle to the next program/item
    Case WM_RBUTTONDOWN
        'If the selection is 5 then we restart at 0, this is
        '5 because that is how many items I chose to add, you
        'could add more or less, just alter the numbers accordingly.
        If CurrentSelection = 5 Then CurrentSelection = 0
            'Increment the selection each time we right click
            CurrentSelection = CurrentSelection + 1
            'Call the SelectionState Sub
            SelectionState
            'Alter our icon
            LoadIcons False
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_Terminate()
'Kill the systray icon
Shell_NotifyIcon NIM_DELETE, icoDat
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Kill the systray icon
Shell_NotifyIcon NIM_DELETE, icoDat
End Sub

Private Sub tmrOnScreen_Timer()
'This timer is used to control the length of time
'the form is displayed for the various program/items.
Me.Hide
tmrOnScreen.Enabled = False
End Sub
