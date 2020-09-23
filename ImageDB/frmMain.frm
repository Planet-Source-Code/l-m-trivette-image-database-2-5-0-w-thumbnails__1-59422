VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   " Image Database"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timResize 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   5640
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10821
      View            =   2
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dimesions"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   1323
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   7680
      TabIndex        =   1
      Top             =   6360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6300
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7223
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3598
            MinWidth        =   3598
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   8040
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picPlane 
      BorderStyle     =   0  'None
      Height          =   6105
      Left            =   0
      ScaleHeight     =   6105
      ScaleWidth      =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1800
      Begin VB.PictureBox picHandle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   1680
         MousePointer    =   9  'Size W E
         ScaleHeight     =   6135
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   0
         Width           =   100
      End
      Begin VB.ListBox List1 
         Height          =   6105
         Left            =   120
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File "
      Begin VB.Menu mnuNew 
         Caption         =   "New Category"
      End
      Begin VB.Menu sep103 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu sep101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu sep100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit "
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy "
         Shortcut        =   ^C
      End
      Begin VB.Menu sep201 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "Invert Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuThumbs 
         Caption         =   "Thumbnails"
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Icons"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Details"
      End
      Begin VB.Menu mnuList 
         Caption         =   "List"
      End
      Begin VB.Menu sep301 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "Arrange By"
         Begin VB.Menu mnuArrangeName 
            Caption         =   "Title"
         End
         Begin VB.Menu mnuArrangeSize 
            Caption         =   "File Size"
         End
         Begin VB.Menu mnuArrangeType 
            Caption         =   "Type"
         End
         Begin VB.Menu mnuArrangeModified 
            Caption         =   "Modified"
         End
      End
      Begin VB.Menu sep302 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools "
      Visible         =   0   'False
      Begin VB.Menu mnuAutoRename 
         Caption         =   "Automatic Renaming"
      End
      Begin VB.Menu mnuSlideShow 
         Caption         =   "Slide Show"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuVote 
         Caption         =   "Vote at Planet Source Code"
         Visible         =   0   'False
      End
      Begin VB.Menu sep501 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuClipboard 
         Caption         =   "Copy to Clipboard"
      End
      Begin VB.Menu sep609 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu sep102 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll2 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuInvertSlection2 
         Caption         =   "Invert Selection"
      End
      Begin VB.Menu mnuRefresh5 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Image Database Version 2.5.x
'
' Written by L. "Mike" Trivette
' Please send me comments at mtrivette@yahoo.com
'
' I'm sorry for any code snippets i used and forgot to give credit.
' I worked hard on this so please give feedback and credit if due.
'
' Last Revised 3/11/2005 12:37:10pm
'
' See Readme file for more information... >>> readme.txt
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public bCancel As Boolean 'Interupts the loadtitles sub
Dim lastindex As Long
Dim ask As Boolean
Dim response As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Cancel the loadtitle loop if user press ESC button
    If KeyCode = 27 Then bCancel = True
End Sub

Private Sub Form_Load()
    ' Show the current version on the form caption
    Me.Caption = " Image Database " & App.Major & "." & App.Minor & "." & App.Revision
    AddProgBar ProgressBar1, StatusBar1, 4 ' Initialize the progressbar into the statusbar
    loadset ' Load the program settings
    loadcat ' Load image categories
    Me.Show ' Show form
End Sub

Private Sub Form_Resize()
    ' Make sure form is not minimized
    If frmMain.WindowState = 1 Then Exit Sub
    ' Make sure form is not resized too small
    If frmMain.Width < 3650 Then frmMain.Width = 3650
    If frmMain.Height < 3650 Then frmMain.Height = 3650
    ' Exit sub if window was minimized
    If frmMain.WindowState = 1 Then Exit Sub
    ' Update the progressbar into the statusbar
    AddProgBar ProgressBar1, StatusBar1, 4
    ' Reposition and refresh the listview
    ListView1.Width = Me.Width - ListView1.Left - 225
    ListView1.Height = Me.Height - StatusBar1.Height - 900
    ' Update the listview control
    ListView1.Arrange = lvwAutoTop
    ListView1.Refresh
    ' Resize the list control based on the size of the listview control
    picPlane.Height = ListView1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Compact the working database if user selected
    If AutoCompact = True Then compressDB
    ' Basic variable declaration
    Dim frm As Form
    ' In case the loadtitle loop is in progress
    bCancel = True
    ' Put all the INI configuration saves here to remeber GUI the settings
    SetValue App.path & "\Config.ini", "Settings", "ListView", ListView1.View
    ' Free up any memory used by any remaining loaded forms
     For Each frm In Forms
          Unload frm
          Set frm = Nothing
     Next frm
End Sub

Private Sub List1_Click()
    bCancel = True
    bCancel = False
    loadtitles
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Cancel the loadtitle loop is user press ESC button
    If KeyCode = 27 Then bCancel = True
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xx As Variant
    xx = Data.GetData(1)
    Me.Caption = xx
End Sub

Private Sub ListView1_Click()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    StatusBar1.Panels(1).Text = ListView1.SelectedItem.Text
    StatusBar1.Panels(2).Text = ListView1.SelectedItem.SubItems(2) & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    StatusBar1.Panels(3).Text = ListView1.SelectedItem.Index & " of " & ListView1.ListItems.Count & " images"
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When a ColumnHeader object is clicked,
    ' the ListView control is sorted by the
    ' subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    Dim icol As Integer
    If ColumnHeader.Index - 1 <> icol Then
        ListView1.SortOrder = 0
    Else
        ListView1.SortOrder = Abs(ListView1.SortOrder - 1)
    End If
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    ListView1.Sorted = True
    icol = ColumnHeader.Index - 1
End Sub

Private Sub ListView1_DblClick()
    ' Basic variable declaration
    Dim Preview As frmPreview
    ' Show multple previews or just single
    If MultiPreview = True Then
        Set Preview = New frmPreview
    Else
        Set Preview = frmPreview
    End If
    ' Write the image to disk temporarily
    WriteImage App.path & "\" & ListView1.SelectedItem.Text
    ' Load the preview image to the preview form
    Preview.Caption = ListView1.SelectedItem.Text
    Preview.Picture1.Picture = LoadPicture(App.path & "\" & ListView1.SelectedItem.Text)
    ' Erase the temporary image file from the disk
    Kill App.path & "\" & ListView1.SelectedItem.Text
    ' Show the preview form
    Preview.Show
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Cancel the loadtitle loop is user press ESC button
    If KeyCode = 27 Then bCancel = True
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Delete image if user hits the del key
    If KeyCode = 46 Then DelImage ListView1.SelectedItem.Text
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If ListView1.ListItems.Count = 0 Then Exit Sub
    ' Show the popup menu if listview control right clicked
    If Button = 2 Then
        LeftClick ' Click on the item under the mouse
        TimeOut (0.1) ' Wait for system to catch up
        If ListView1.SelectedItem.Text = "" Then Exit Sub
        PopupMenu frmMain.mnuHidden ' Show popup menu
    End If
End Sub

Private Sub mnuAbout_Click()
    ' Show the Spash/About screen
    frmSplash.Show
End Sub

Private Sub mnuArrangeModified_Click()
    ClearArranges
    mnuArrangeModified.Checked = True
End Sub

Private Sub mnuArrangeName_Click()
    ClearArranges
    mnuArrangeName.Checked = True
End Sub

Private Sub mnuArrangeSize_Click()
    ClearArranges
    mnuArrangeSize.Checked = True
End Sub

Private Sub mnuArrangeType_Click()
    ClearArranges
    mnuArrangeType.Checked = True
End Sub

Private Sub mnuAutoRename_Click()
    frmAutoRename.Show
End Sub

Private Sub mnuClipboard_Click()
        
    ' Load selected image into picture control
    WriteImage App.path & "\" & ListView1.SelectedItem.Text
    Picture1.Picture = LoadPicture(App.path & "\" & ListView1.SelectedItem.Text)
    Kill App.path & "\" & ListView1.SelectedItem.Text
    
    ' Make sure image type is clipboard compatible
    If Picture1.Picture.Type <> 1 Then
        response = MsgBox("Cannot copy selected image to clipboard." & vbCrLf & vbCrLf & "Wrong image type.", vbOKOnly + vbExclamation, " Clipboard Error")
        Exit Sub
    End If
    
    ' Set image to clipboard
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture
    
    ' Update the status bar to notify user the task was completed
    StatusBar1.Panels(1).Text = ListView1.SelectedItem.Text & " copied to clipboard"
End Sub

Private Sub mnuDelete_Click()
    Dim i As Long
    
    For i = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(i).Selected = True Then
            DelImage ListView1.ListItems(i)
            ListView1.ListItems.Remove (i)
        End If
    Next i
    ListView1.Arrange = lvwAutoTop
    ListView1.Refresh

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExport_Click()
    Dim response As String
    Const CDERR_CANCEL = &H7FF3

    If ListView1.ListItems.Count = 0 Then Exit Sub

    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Save file"
        .InitDir = GetValue(App.path & "\Config.ini", "Settings", "ExportDir", CurDir)
        .Filename = ListView1.SelectedItem.Text
        .DefaultExt = "." & ListView1.SelectedItem.SubItems(3)
        
        ' Make sure cancel button was not pressed
        If Err = CDERR_CANCEL Then
            MsgBox "mother fucker"
            Exit Sub
        End If
    
        .ShowSave
    End With
    
    ' Write the image from the database
    WriteImage CommonDialog1.Filename
    
    ' Save the last directory used so you can go there automatically next time
    SetValue App.path & "\Config.ini", "Settings", "ExportDir", CurDir
    
    ' Update the statusbar to notify user
    StatusBar1.Panels(1).Text = "Saved " & ListView1.SelectedItem.Text
    
End Sub

Private Sub mnuImport_Click()
    ' Many thanks and credit goto Pietro Cecchi for debugging this sub for me.
    
    Dim BufferFileArray() As String
    Dim i As Integer
    Dim Token As Long
    
    ' Initialise GDI+
    Token = InitGDIPlus
    With CommonDialog1
        .DialogTitle = "Add Multiple files..."
        .Filter = "All Image Files|*.jpg;*.jpeg;*.gif;*.bmp;*.ico;*.w mf"
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
        .InitDir = GetValue(App.path & "\Config.ini", "Settings", "ImportDir", CurDir)
        .MaxFileSize = 32768 - 1
        '32KB=32*1024=32768,default 256, augmented because multiselect (many names)
        .Filename = ""
        .ShowOpen
        BufferFileArray = Split(.Filename, Chr(0))
        'where BufferFileArray(0) is the path, and following BufferFileArray(i) are the images names
    End With
    
    ' If no files are selected
    If UBound(BufferFileArray) = -1 Then GoTo exitsublabel
    
    ' If only one file was chosen.
    If UBound(BufferFileArray) = 0 Then
        SaveImage CommonDialog1.Filename
        Exit Sub
    End If
    
    ' If multiple files chosen.
    ProgressBar1.Max = UBound(BufferFileArray)
    For i = LBound(BufferFileArray) + 1 To UBound(BufferFileArray)
        ProgressBar1.Value = i
        SaveImage BufferFileArray(0) & "\" & BufferFileArray(i)
    Next
    
    ' Reset
    ProgressBar1.Value = 0
    
exitsublabel:
    ' Free GDI+
    FreeGDIPlus Token
   
    ' Save the last directory used so you can go there automatically next time
    SetValue App.path & "\Config.ini", "Settings", "ImportDir", CurDir
    
End Sub

Private Sub mnuInvert_Click()
    Dim i As Long
    If ListView1.ListItems.Count = 0 Then Exit Sub
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            ListView1.ListItems(i).Selected = False
        Else
            ListView1.ListItems(i).Selected = True
        End If
    Next i
End Sub

Private Sub mnuInvertSlection2_Click()
    mnuInvert_Click
End Sub

Private Sub mnuNew_Click()
    Dim response As String
    
    response = InputBox("Name of new category" & vbCrLf & vbCrLf & "Note: If you do not store any images into the new category then the category will not be saved.", "New Category", "")
    If response = "" Then Exit Sub
    
    List1.AddItem response
    List1.Selected(List1.ListCount - 1) = True
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuPreview_Click()
    ListView1_DblClick
End Sub

Private Sub mnuRefresh_Click()
    bCancel = True
    bCancel = False
    loadtitles
End Sub

Private Sub mnuRefresh5_Click()
    ListView1.Arrange = lvwAutoTop
    ListView1.Refresh
End Sub

Private Sub mnuRename_Click()
    RenameImage ListView1.SelectedItem.Text
End Sub

Private Sub mnuSaveAs_Click()
    mnuExport_Click
End Sub

Private Sub mnuSelectAll_Click()
    Dim i As Long
    If ListView1.ListItems.Count = 0 Then Exit Sub
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Selected = True
    Next i
End Sub

Private Sub mnuSelectAll2_Click()
    mnuSelectAll_Click
End Sub

Private Sub mnuSlideShow_Click()
    If ListView1.ListItems.Count > 0 Then frmSlideShow.Show
End Sub

Private Sub mnuThumbs_Click()
    ListView1.View = lvwIcon
    ClearViews
    mnuThumbs.Checked = True
    bCancel = True
    bCancel = False
    loadtitles
End Sub

Private Sub mnuDetails_Click()
    ListView1.View = lvwReport
    ClearViews
    mnuDetails.Checked = True

End Sub

Private Sub mnuIcons_Click()
    ListView1.View = lvwSmallIcon
    ClearViews
    mnuIcons.Checked = True

End Sub

Private Sub mnuList_Click()
    ListView1.View = lvwList
    ClearViews
    mnuList.Checked = True

End Sub

Private Sub ClearViews()
    mnuList.Checked = False
    mnuIcons.Checked = False
    mnuDetails.Checked = False
    mnuThumbs.Checked = False
End Sub

Private Sub ClearArranges()
    mnuArrangeModified.Checked = False
    mnuArrangeName.Checked = False
    mnuArrangeSize.Checked = False
    mnuArrangeType.Checked = False
End Sub

Public Sub loadtitles()
    ' Assign variables
    Dim dbs As Database
    Dim rst As Recordset
    Dim itmx As ListItem
    Dim strFile As String
    Dim strsql As String
    Dim Token As Long
    Dim i As Long
    
    ' Clear the controls
    Set ListView1.Icons = Nothing
    ListView1.ListItems.Clear
    ImageList1.ListImages.Clear
    
    ' Update statusbar
    StatusBar1.Panels(1).Text = "Loading images from database... Press ESC to Cancel"
    StatusBar1.Panels(2).Text = ""
    
    '
    If List1.SelCount = 0 Or List1.Text = "All images" Then
        List1.Selected(0) = True
        strsql = "SELECT * from images;"
    Else
        strsql = "SELECT * FROM images where category = '" & List1.Text & "';"
    End If
    
    ' Initialise GDI+
    Token = InitGDIPlus
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset(strsql)
    If rst.RecordCount > 0 Then
        mnuTools.Visible = True
        rst.MoveLast: rst.MoveFirst
        ProgressBar1.Max = rst.RecordCount

        For i = 1 To rst.RecordCount
            If bCancel = True Then GoTo skip
            DoEvents
            ProgressBar1.Max = rst.RecordCount
            ProgressBar1.Value = i
            StatusBar1.Panels(3).Text = "  Loading " & i & " of " & ProgressBar1.Max & "  "
            strFile = rst.Fields("title")
                        
            If ListView1.View = 0 Then
                WriteImage App.path & "\" & strFile
                ImageList1.ListImages.Add , , modGDIPlusResize.LoadPictureGDIPlus(App.path & "\" & strFile, ThumbWidth, ThumbHeight, , True)
                Set ListView1.Icons = ImageList1
                Set itmx = ListView1.ListItems.Add(, , strFile, ImageList1.ListImages.Count)
            Else
                Set itmx = ListView1.ListItems.Add(, , strFile)
            End If
                itmx.SubItems(1) = SetBytes(rst.Fields("size"))
                itmx.SubItems(2) = rst.Fields("width") & " x " & rst.Fields("height")
                itmx.SubItems(3) = rst.Fields("type")
            
            If ListView1.View = 0 Then
                Kill App.path & "\" & strFile
            End If
            
            rst.MoveNext
            
            
        Next i
        bCancel = True
skip:
        StatusBar1.Panels(3).Text = ListView1.ListItems.Count & " images"
        ProgressBar1.Value = 0
    End If
    rst.Close
    dbs.Close
    
    StatusBar1.Panels(1).Text = ""
    
    ' Free GDI+
    FreeGDIPlus Token
    
    ListView1.Arrange = lvwAutoTop
    ListView1.Refresh
End Sub

Private Sub loadset()
    ' Load all the program settings from the configuration file.
    srcDB = GetValue(App.path & "\Config.ini", "Settings", "Database", App.path & "\images.mdb")
    ThumbWidth = GetValue(App.path & "\Config.ini", "Settings", "ThumbWidth", 100)
    ThumbHeight = GetValue(App.path & "\Config.ini", "Settings", "ThumbHeight", 100)
    ExportPath = GetValue(App.path & "\Config.ini", "Settings", "ExportPath", App.path & "\")
    DelImport = GetValue(App.path & "\Config.ini", "Settings", "DelImport", False)
    DelExport = GetValue(App.path & "\Config.ini", "Settings", "DelExport", False)
    MultiPreview = GetValue(App.path & "\Config.ini", "Settings", "MultiPreview", False)
    AutoCompact = GetValue(App.path & "\Config.ini", "Settings", "AutoCompact", True)
        
    ' Load the GUI preferences
    ListView1.View = GetValue(App.path & "\Config.ini", "Settings", "ListView", lvwIcon)

    ' Set the "VIEW" menu settings to match settings
    Select Case ListView1.View
    Case lvwIcon
        mnuThumbs.Checked = True
    Case lvwList
        mnuList.Checked = True
    Case lvwSmallIcon
        mnuIcons.Checked = True
    Case lvwReport
        mnuDetails.Checked = True
    End Select
    
End Sub

Public Sub loadcat()
    ' Assign Variables
    Dim dbs As Database
    Dim rst As Recordset
    Dim i As Long
    ' Clear the list control
    List1.Clear
    ' Open Datbase
    Set dbs = OpenDatabase(srcDB)
    ' Get recordset from opened database
    Set rst = dbs.OpenRecordset("SELECT category FROM images group by category;")
    ' If no images with categories were found
    If rst.RecordCount = 0 Then GoTo skip
    ' Populate the recordset and set to beginning
    rst.MoveLast: rst.MoveFirst
    ' Add the "All images" category to the list control
    List1.AddItem "All images"
    'Cycle throught the recordset and add to list control
    For i = 1 To rst.RecordCount
        List1.AddItem rst.Fields("category")
        rst.MoveNext
    Next i
skip:
    ' Close the recordset and database
    rst.Close
    dbs.Close
End Sub


Private Sub picHandle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*' The MouseDown event in the picHandle object is the trigger that will allow the tracking
    '*' of the mouse position to begin.  All of the calculations are done in a timer to allow
    '*' for the repetative and constant tracking of the mouse position and object sizes.
    '*'
    timResize.Enabled = True
End Sub

Private Sub picHandle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*' The MouseUp event kills the trigger that was set on the MouseDown event.  This will mean
    '*' that the user has release the mouse button and does not wish to resize the split plane
    '*' any further.
    '*'
    timResize.Enabled = False
End Sub


Private Sub picPlane_Resize()
    List1.Width = picPlane.Width - 225
    List1.Height = picPlane.Height + 25
    picHandle.Left = picPlane.Width - picHandle.Width
    
    
    
End Sub

Private Sub timResize_Timer()
    Dim MinWidth As Long            '*' Minimum Width of the Split Plane
    Dim MaxWidth As Long            '*' Maximum Width of the Split Plane
    Dim lngRelX As Long             '*' Calculated value of the width of the Split Plane
    Dim CurrentX As Long            '*' Current XValue of the Mouse
    Dim intcalc As Long             '*'
    Static LastMousePosX As Long    '*' Static Variable to Track the Mouse's Last Known Position
    
    '*' Get the current value of the X based upon the location of the mouse.
    '*'
    CurrentX = GetX
    
    '*' If the value has not changed since the last time, the user has not moved position.
    '*' Exit the subroutine to skip redundant calculations.
    '*'
    If CurrentX = LastMousePosX Then
        Exit Sub
    Else
        '*' If the value is different, set the value to the current value of X for the next time
        '*' this event fires.
        '*'
        LastMousePosX = CurrentX
    End If
    
    '*' The minimum and maximum width of the splitter plane can be set to either an absolute value
    '*' or to an equation.  Values are represented in Twips.
    '*'
    MinWidth = 1000                 '*' On some machines, smaller values cause jumpiness.
    MaxWidth = 7450   '*' Limit the maximum to be one half of the form's size.
  
    intcalc = (CurrentX * Screen.TwipsPerPixelX) - Me.Left
          
    '*' Bounds Checking.  Make sure that the value returned is within bounds.  If it is not, set
    '*' it to the proper value and exit the sub.
    '*'
    If intcalc <= MinWidth Then
      intcalc = MinWidth
    ElseIf intcalc >= MaxWidth Then
      intcalc = MaxWidth
    End If
    
    '*' Set the width of the Split Plane to be equal to that
    picPlane.Width = intcalc ' Me.Left + (CurrentX * (picPlane.Width / picPlane.ScaleWidth))

    ListView1.Left = intcalc
    ListView1.Width = Me.Width - ListView1.Left - 225
    
End Sub
