VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Options"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImageDatabase.LaVolpeButton LaVolpeButton6 
      Height          =   495
      Left            =   1080
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Apply"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmOptions.frx":0ECA
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin ImageDatabase.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmOptions.frx":0EE6
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Database "
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5535
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto compact on exit"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   1400
         Width           =   1935
      End
      Begin ImageDatabase.LaVolpeButton LaVolpeButton5 
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "Compact Database"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmOptions.frx":0F02
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin ImageDatabase.LaVolpeButton LaVolpeButton4 
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "Select Database"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmOptions.frx":0F1E
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin ImageDatabase.LaVolpeButton LaVolpeButton3 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "Create New Database"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmOptions.frx":0F3A
         ALIGN           =   1
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
         OPTVAL          =   0   'False
         OPTMOD          =   0   'False
         GStart          =   0
         GStop           =   16711680
         GStyle          =   0
      End
      Begin VB.Label lblDatabase 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "C:\"
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   275
         Width           =   5295
      End
   End
   Begin ImageDatabase.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmOptions.frx":0F56
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
      Begin VB.CheckBox Check5 
         Caption         =   "Cache thumbnails"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete Images after exporting"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Delete Images after importing"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show multiple previews"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Thumbnail Size"
      Height          =   1935
      Left            =   3240
      TabIndex        =   0
      Top             =   2040
      Width           =   2415
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Symmetric"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   285
         LargeChange     =   10
         Left            =   1800
         Max             =   500
         Min             =   20
         TabIndex        =   6
         Top             =   900
         Value           =   20
         Width           =   185
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   285
         LargeChange     =   10
         Left            =   1800
         Max             =   500
         Min             =   20
         TabIndex        =   3
         Top             =   480
         Value           =   20
         Width           =   185
      End
      Begin VB.Label Label4 
         Caption         =   "Height:"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Width:"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check3_Click()
    If Check3.Value = 1 Then Text4.Text = Text3.Text
End Sub

Private Sub Command1_Click()

End Sub



Private Sub Form_Load()
    ' Set the form controls based on the global variables
    lblDatabase.Caption = srcDB
    Text3.Text = ThumbWidth
    Text4.Text = ThumbHeight
    VScroll1.Value = ThumbWidth
    VScroll2.Value = ThumbHeight
    
    
    If DelExport = True Then Check1.Value = 1
    If DelImport = True Then Check2.Value = 1
    If MultiPreview = True Then Check4.Value = 1
    If AutoCompact = True Then Check6.Value = 1
    
    ' Load this from directly form the config.ini since it is not a global variable
    Check3.Value = GetValue(App.path & "\Config.ini", "Settings", "Symmetric", 0)
    
End Sub

Private Sub LaVolpeButton1_Click()
    saveset
    Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
    Unload Me
End Sub

Private Sub LaVolpeButton3_Click()
    With frmMain.CommonDialog1
        .DialogTitle = "Select new database"
        .Filter = "Image Database|*.mdb"
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
        .Filename = ""
        .ShowSave
    End With
    
    If frmMain.CommonDialog1.Filename = "" Then Exit Sub
    
    FileCopy App.path & "\shell", frmMain.CommonDialog1.Filename

    If TestDb(frmMain.CommonDialog1.Filename) = True Then
        srcDB = frmMain.CommonDialog1.Filename
        lblDatabase.Caption = srcDB
        frmMain.loadcat
        frmMain.ListView1.ListItems.Clear
    End If
    
End Sub

Private Sub LaVolpeButton4_Click()
    With frmMain.CommonDialog1
        .DialogTitle = "Select new database"
        .Filter = "Image Database|*.mdb"
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
        .Filename = ""
        .ShowOpen
    End With
    
    If frmMain.CommonDialog1.Filename = "" Then Exit Sub
    
    If TestDb(frmMain.CommonDialog1.Filename) = True Then
        srcDB = frmMain.CommonDialog1.Filename
        lblDatabase.Caption = srcDB
        frmMain.loadcat
        frmMain.ListView1.ListItems.Clear
    End If
    

End Sub

Private Sub LaVolpeButton5_Click()
    Dim before As String
    Dim after As String
    Dim response As String
    
    before = GetFileSize(srcDB)
    compressDB
    after = GetFileSize(srcDB)
    
    response = MsgBox("Database was successfully compacted.                               " & vbCrLf & vbCrLf & "Database filename: " & srcDB & vbCrLf & vbCrLf & "      Filesize before :  " & before & vbCrLf & "      Filesize after :  " & after & vbCrLf & vbCrLf, vbOKOnly, " Compact Database")
    
End Sub

Private Sub LaVolpeButton6_Click()
    saveset
End Sub

Private Sub Text3_Change()
    If IsNumeric(Text3.Text) = True And Check3.Value = 1 Then Text4.Text = Text3.Text
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 9 And KeyAscii <> 32 And KeyAscii <> 8 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub

Private Sub Text3_LostFocus()
    ' Do error checking on thumbnail width
    If Text3.Text = "" Then Text3.Text = ThumbWidth
    If Val(Text3.Text) > 500 Then
        Text3.Text = "500"
        VScroll1.Value = 500
    End If
    VScroll1.Value = Val(Text3.Text)
    If Check3.Value = 1 Then VScroll2.Value = Val(Text3.Text)
End Sub

Private Sub Text4_Change()
    If IsNumeric(Text4.Text) = True And Check3.Value = 1 Then Text3.Text = Text4.Text
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 9 And KeyAscii <> 32 And KeyAscii <> 8 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub

Private Sub Text4_LostFocus()
    ' Do error checking on thumbnail width
    If Text4.Text = "" Then Text4.Text = ThumbHeight
    If Val(Text4.Text) > 500 Then
        Text4.Text = "500"
        VScroll2.Value = 500
    End If
    VScroll2.Value = Val(Text4.Text)
    If Check3.Value = 1 Then VScroll1.Value = Val(Text4.Text)
End Sub

Private Sub VScroll1_Change()
    Text3.Text = VScroll1.Value
    If Check3.Value = 1 Then Text4.Text = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
    Text4.Text = VScroll2.Value
    If Check3.Value = 1 Then Text3.Text = VScroll2.Value
End Sub

Private Sub saveset()
        ' Save the settings to config.ini file
    SetValue App.path & "\Config.ini", "Settings", "Database", lblDatabase.Caption
    SetValue App.path & "\Config.ini", "Settings", "DelExport", CBool(Check1.Value)
    SetValue App.path & "\Config.ini", "Settings", "DelImport", CBool(Check2.Value)
    SetValue App.path & "\Config.ini", "Settings", "MultiPreview", CBool(Check4.Value)
    SetValue App.path & "\Config.ini", "Settings", "AutoCompact", CBool(Check6.Value)
    SetValue App.path & "\Config.ini", "Settings", "Symmetric", Check3.Value
    SetValue App.path & "\Config.ini", "Settings", "ThumbWidth", Text3.Text
    SetValue App.path & "\Config.ini", "Settings", "ThumbHeight", Text4.Text
    
    ' Reset the program global variables
    srcDB = lblDatabase.Caption
    ThumbWidth = Text3.Text
    ThumbHeight = Text4.Text
    DelExport = CBool(Check1.Value)
    DelImport = CBool(Check2.Value)
    MultiPreview = CBool(Check4.Value)
    AutoCompact = CBool(Check6.Value)
End Sub
