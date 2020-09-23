Attribute VB_Name = "modDatabaseFunc"
Const Blocksize = 32768

' Program wide variables
Global srcDB As String
Global ExportPath As String
Global ThumbWidth As Long
Global ThumbHeight As Long
Global DelImport As Boolean
Global DelExport As Boolean
Global MultiPreview As Boolean
Global AutoCompact As Boolean


Public Sub SaveImage(strImage As String)
    On Error GoTo skip
    ' Save image to database
    Dim NumBlocks As Integer, SourceFile As Integer, i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim FileData() As Byte, retval As Variant
    Dim dbs As Database
    Dim rst As Recordset
    Dim strHex As String
    Dim itmx As ListItem
    
    Set m_CRC = New clsCRC
    
    frmMain.Picture1.Cls
    frmMain.Picture1.Picture = LoadPicture(strImage)
    
    frmMain.StatusBar1.Panels(1).Text = "Importing " & GetFileName(Replace(strImage, "'", ""))
    
    strHex = Hex(m_CRC.CalculateFile(strImage))
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM images where crc = '" & strHex & "';")
           
    m_CRC.Algorithm = CRC32
    
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("title") = GetFileName(Replace(strImage, "'", ""))
        rst.Fields("crc") = strHex
        rst.Fields("size") = FileLen(strImage)
        rst.Fields("width") = frmMain.Picture1.Width / 15
        rst.Fields("height") = frmMain.Picture1.Height / 15
        rst.Fields("type") = LCase(GetFileExtension(strImage))
        If frmMain.List1.Text = "" Or frmMain.List1.Text = "All images" Then
            rst.Fields("category") = "unsorted"
        Else
            rst.Fields("category") = frmMain.List1.Text
        End If
        
        SourceFile = FreeFile
        Open strImage For Binary Access Read As SourceFile
        FileLength = LOF(SourceFile)
            NumBlocks = FileLength \ Blocksize
            LeftOver = FileLength Mod Blocksize 'remainder appended first
            ReDim FileData(LeftOver)
            Get SourceFile, , FileData()
            rst.Fields("BinData").AppendChunk FileData() 'store the first image chunk
            ReDim FileData(Blocksize)
            For i = 1 To NumBlocks
                Get SourceFile, , FileData()
                rst.Fields("BinData").AppendChunk FileData() 'remaining chunks
                DoEvents
            Next i
        Close SourceFile
        rst.Update
    Else
        ' if duplicate image found
        If ask = True Then response = MsgBox("This image was already found in database." & vbCrLf & vbCrLf & "Source: " & GetFileName(Replace(strImage, "'", "")) & vbCrLf & "Found: " & rst.Fields("title") & vbCrLf & vbCrLf & "Would you like to continue to receive duplicate warnings?", vbYesNo + vbInformation, "Duplicate")
        If response = 7 Or response = 0 Then
            ask = False
        Else
            ask = True
        End If
    End If
    
    rst.Close
    dbs.Close
    
    If frmMain.ListView1.View = 0 Then
         frmMain.ImageList1.ListImages.Add , , modGDIPlusResize.LoadPictureGDIPlus(strImage, ThumbWidth, ThumbHeight, , True)
         Set frmMain.ListView1.Icons = frmMain.ImageList1
         Set itmx = frmMain.ListView1.ListItems.Add(, , GetFileName(strImage), frmMain.ImageList1.ListImages.Count)
    Else
         Set itmx = frmMain.ListView1.ListItems.Add(, , GetFileName(strImage))
    End If
            
         itmx.SubItems(1) = FileLen(strImage)
         itmx.SubItems(2) = frmMain.Picture1.Width / 15 & " x " & frmMain.Picture1.Height / 15
         itmx.SubItems(3) = LCase(GetFileExtension(strImage))
                
    ' Delete the source file if user wants
    If DelImport = True Then Kill strImage
    frmMain.StatusBar1.Panels(2).Text = ""
    frmMain.StatusBar1.Panels(3).Text = frmMain.ListView1.ListItems.Count & " images"
    
skip:
End Sub

Public Sub DelImage(strImage As String)
    Dim dbs As Database
    Dim rst As Recordset
    Dim strTitle As String
        
    frmMain.StatusBar1.Panels(2).Text = "Deleting " & strImage
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("Select * from images where title = '" & strImage & "';")
    ' Delete image if found
    If rst.RecordCount = 0 Then
        MsgBox "Error deleting " & strImage & " from database."
    Else
        rst.Delete
    End If
    rst.Close
    dbs.Close
    
    frmMain.StatusBar1.Panels(2).Text = ""

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get image from database and write to disk                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WriteImage(strImage As String, Optional strfolder As String)
    ' Assign variables
    Dim dbs As Database
    Dim rst As Recordset
    Dim response As String
    ' Open the datbase and then recordset containing the images
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM images where title = '" & GetFileName(strImage) & "';")
    ' Write image to disk if found in recordset
    If rst.RecordCount > 0 Then
        Call WriteBLOB(rst, "BinData", strImage)
        frmMain.Picture1.Picture = LoadPicture(strImage)
        frmMain.Picture1.Refresh
    Else
        response = MsgBox(strImage & " not found in database.", vbOKOnly + vbExclamation, "Error")
    End If
    ' Close recordset and database
    rst.Close
    dbs.Close
    ' Release variables
    Set rst = Nothing
    Set dbs = Nothing
End Sub

'**********************************************************************
'FUNCTION: WriteBLOB()
'
'PURPOSE:
'WritesBLOB information stored in the specified table and field to the
'specified disk file.
'
'PREREQUISITES:
'
'ARGUMENTS:
'Destination - the path and filename of the file to be extracted.
'T - the table object the data is stored in.
'Field - the OLE object to store the data in.
'
'RETURN:
'0 on fail 1 on success
'**********************************************************************
Public Function WriteBLOB(T As Recordset, sField As String, Destination As String)
        On Error GoTo Err_WriteBLOB
        Dim NumBlocks As Integer, DestFile As Integer, i As Integer
        Dim FileLength As Long, LeftOver As Long
        Dim FileData() As Byte, retval As Variant

        ' Get the length of the file.
        FileLength = T(sField).FieldSize()
        If FileLength <> 0 Then
            DestFile = FreeFile
            NumBlocks = FileLength \ Blocksize
            LeftOver = FileLength Mod Blocksize 'reminder appended first
            'initialize status bar meter
            'RetVal = SysCmd(acSysCmdInitMeter, "Writing BLOB", NumBlocks)

            Open Destination For Binary Access Write Lock Write As DestFile
            ReDim FileData(LeftOver)
            FileData() = T(sField).GetChunk(0, LeftOver)
            Put DestFile, , FileData() 'write first chunk
            
            ReDim FileData(Blocksize)
            
            For i = 1 To NumBlocks
                FileData() = T(sField).GetChunk((i - 1) * Blocksize + LeftOver, Blocksize)
                Put DestFile, , FileData() 'write remaining chunks
                'update status bar meter
                'RetVal = SysCmd(acSysCmdUpdateMeter, i)
            Next i
            Close DestFile
        End If
        
        'remove status bar meter
        'RetVal = SysCmd(acSysCmdRemoveMeter)
        WriteBLOB = 1
        Exit Function

Err_WriteBLOB:
        MsgBox Err.Description
        WriteBLOB = 0
        Exit Function
End Function

Public Function RenameImage(strImage As String)
    ' Todo: Do some error checking on the response from user
    '     : ensure file extension integrity
    
    Dim dbs As Database
    Dim rst As Recordset
    Dim strTitle As String
    Dim response As String
    
    frmMain.StatusBar1.Panels(2).Text = "Renaming " & strImage
    
    response = InputBox("Rename: " & vbCrLf & vbCrLf & strImage, "Rename", "")
    If response = "" Then Exit Function
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("Select * from images where title = '" & strImage & "';")
    ' Rename image if found
    If rst.RecordCount = 0 Then
        MsgBox "Error deleting " & strImage & " from database."

    Else
        rst.Edit
        rst.Fields("title") = response
        rst.Update
        frmMain.ListView1.ListItems(frmMain.ListView1.SelectedItem.Index).Text = response
        frmMain.ListView1.ListItems(frmMain.ListView1.SelectedItem.Index).Selected = True
    End If
    rst.Close
    dbs.Close
    
    frmMain.StatusBar1.Panels(2).Text = ""


End Function

Public Function TestDb(srcTestDB As String) As Boolean
    On Error GoTo skip
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = OpenDatabase(srcTestDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM images;")
    
    rst.Close
    dbs.Close
    
    Set rst = Nothing
    Set dbs = Nothing
    
    TestDb = True
    Exit Function
skip:
    TestDb = Flase
End Function




Public Sub compressDB()
    frmMain.StatusBar1.Panels(1).Text = "Compacting Database..."
    CompactDatabase srcDB, "temp_0000.mdb"
    If TestDb("temp_0000.mdb") = True Then
        frmMain.StatusBar1.Panels(1).Text = "Compacting Database...   Writing database."
        Kill srcDB
        FileCopy "temp_0000.mdb", srcDB
        frmMain.StatusBar1.Panels(1).Text = "Checking database integrity"
        If TestDb(srcDB) = True Then Kill "temp_0000.mdb"
    End If
    frmMain.StatusBar1.Panels(1).Text = ""
End Sub










