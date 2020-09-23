VERSION 5.00
Begin VB.Form frmSlideShow 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5205
   ClientLeft      =   150
   ClientTop       =   90
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7515
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   6840
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuInterval 
         Caption         =   "Interval"
      End
      Begin VB.Menu sep100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
      End
      Begin VB.Menu sep101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1_Timer
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub Form_Resize()
    Label1.Move 0, Me.Height - 800, Me.Width, 500
End Sub

Private Sub Picture1_Change()
    ' Postion form so that it will be centered
    Picture1.Top = (Screen.Height * 0.85) / 2 - Picture1.Height / 2 + 350
    Picture1.Left = Screen.Width / 2 - Picture1.Width / 2
    ' Make sure form top is not off screen
    If Picture1.Top < 50 Then Picture1.Top = 50
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Static curImage As Long
    If frmMain.ListView1.ListItems.Count = 0 Then Unload Me
    
    curImage = curImage + 1
    
    If curImage > frmMain.ListView1.ListItems.Count Then curImage = 1
    
    ' Write the image to disk temporarily
    WriteImage App.path & "\" & frmMain.ListView1.ListItems(curImage).Text
    ' Load the preview image to the preview form
    'Preview.Caption = ListView1.SelectedItem.Text
    Picture1.Picture = LoadPicture(App.path & "\" & frmMain.ListView1.ListItems(curImage).Text)
    ' Erase the temporary image file from the disk
    Kill App.path & "\" & frmMain.ListView1.ListItems(curImage).Text
    
    Label1.Caption = frmMain.ListView1.ListItems(curImage).Text
    
End Sub
