VERSION 5.00
Begin VB.Form frmAutoRename 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Auto Renaming Utility"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmAutoRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin ImageDatabase.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmAutoRename.frx":08CA
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
      Caption         =   "Preview"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   4335
      Begin VB.Label Label4 
         Caption         =   "0001.jpg"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4095
      End
   End
   Begin ImageDatabase.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Start"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "frmAutoRename.frx":08E6
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
      Caption         =   "What images"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton Option1 
         Caption         =   "All Images"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Selected Images"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rename to Match Pattern"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "Remove file extension "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Beginning Number"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Base Name"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAutoRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Check1_Click()
    Label4.Caption = NewTitle
End Sub

Private Sub Form_Load()
    ' If user has chosen any images from the listview then assume selection only
    If CountSelectedItemsInListview(frmMain.ListView1) >= 2 Then
        Option2.Value = True
    Else
        Option1.Value = True
    End If
    Text1.Text = frmMain.List1.Text & "_"
    
    Label4.Caption = Text1.Text & Format(Text2.Text, "0000") & ".jpg"
End Sub

Private Sub LaVolpeButton1_Click()
    If Option1.Value = True Then
        RenameAll
    Else
        RenameSome
    End If
    
    Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
'   Dim itemx As ListItem
'   Dim myCol As Collection
'
'    Set myCol = GetSelectedItemsFromListview(frmMain.ListView1)
'    Text1.Text = ""
'
'    For Each itemx In myCol
'        Text1.Text = Text1.Text & itemx.Text & vbCrLf
'    Next itemx
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Label4.Caption = NewTitle
End Sub

Private Sub Text2_Change()
    ' Just in case user tries to paste in some bullsh*t
    If IsNumeric(Text2) = False Then
        Text2.Text = "1"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Numbers As Integer
    Numbers = KeyAscii
    ' Allow only numbers to entered
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then KeyAscii = 0
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    Label4.Caption = NewTitle
End Sub

Private Function NewTitle() As String
    If Check1.Value = 1 Then
        NewTitle = Text1.Text & Format(Text2.Text, "0000")
    Else
        NewTitle = Text1.Text & Format(Text2.Text, "0000") & ".jpg"
    End If
End Function

Private Sub RenameAll()
    Dim intSeq As Long
    Dim i As Long
    Dim dbs As Database
    Dim rst As Recordset
    Dim newname As String
    
    Me.Hide
    intSeq = Text2.Text
    
    Set dbs = OpenDatabase(srcDB)

    For i = 1 To frmMain.ListView1.ListItems.Count
        Set rst = dbs.OpenRecordset("SELECT * FROM images where title = '" & frmMain.ListView1.ListItems(i).Text & "';")
        If rst.RecordCount > 0 Then
            If Check1.Value = 1 Then
                newname = Text1.Text & Format(intSeq, "0000")
            Else
                newname = Text1.Text & Format(intSeq, "0000") & "." & frmMain.ListView1.ListItems(i).SubItems(3)
            End If
            rst.Edit
            rst.Fields("title") = newname
            rst.Update
            frmMain.ListView1.ListItems(i).Text = newname
        Else
            MsgBox "Could not find to rename"
        End If
        
        intSeq = intSeq + 1
    Next i
    
    rst.Close
    dbs.Close
End Sub

Private Sub RenameSome()
Dim intSeq As Long
    Dim i As Long
    Dim dbs As Database
    Dim rst As Recordset
    Dim newname As String
    
    Me.Hide
    intSeq = Text2.Text
    
    Set dbs = OpenDatabase(srcDB)

    For i = 1 To frmMain.ListView1.ListItems.Count
        
        Set rst = dbs.OpenRecordset("SELECT * FROM images where title = '" & frmMain.ListView1.ListItems(i).Text & "';")
        If rst.RecordCount > 0 And frmMain.ListView1.ListItems(i).Selected = True Then
            If Check1.Value = 1 Then
                newname = Text1.Text & Format(Text2.Text, "0000")
            Else
                newname = Text1.Text & Format(Text2.Text, "0000") & "." & frmMain.ListView1.ListItems(i).SubItems(3)
            End If
            rst.Edit
            rst.Fields("title") = newname
            rst.Update
            frmMain.ListView1.ListItems(i).Text = newname
        End If
        intSeq = intSeq + 1
    Next i
    
    rst.Close
    dbs.Close
End Sub

