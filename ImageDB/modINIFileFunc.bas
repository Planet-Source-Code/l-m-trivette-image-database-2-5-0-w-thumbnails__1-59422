Attribute VB_Name = "modINIFileFunc"
Option Explicit

Private Declare Function getprivateprofilestring Lib "Kernel32" _
 Alias "GetPrivateProfileStringA" (ByVal grpnm As Any, _
 ByVal parnm As Any, ByVal deflt As String, _
 ByVal parvl As String, ByVal parlen As Long, _
 ByVal INIPath As String) As Long

Private Declare Function writeprivateprofilestring Lib "Kernel32" _
 Alias "WritePrivateProfileStringA" (ByVal grpnm As String, _
 ByVal parnm As Any, ByVal parvl As Any, _
 ByVal INIPath As String) As Long

Public Sub DeleteGroup(INIPath As String, INIGroup As String)

' Delete entire group from .INI file (physically)
' Input: inipath=path to .INI file, INIGroup=group to delete
Dim X As Long

X = writeprivateprofilestring(INIGroup, 0&, 0&, INIPath)

End Sub

Public Function GetValues(INIPath As String, INIGroup As String) As Collection

Dim File As Integer
Dim iniName As String
Dim keys As String
Dim finished As Boolean
Dim buffer As String
Dim pos As Integer

  File = FreeFile
  finished = False
  iniName = "[" & INIGroup & "]"

  Open INIPath For Input As File
  Set GetValues = New Collection
  Do
    Line Input #File, buffer
    If Left(buffer, Len(INIGroup) + 2) = iniName Then
      Do
        If Not EOF(File) Then
          Line Input #File, keys
          If Left(keys, 1) = "[" Or keys = "" Then 'check if another key starts or if this one ends or if its EOF
            finished = True
          Else
            pos = InStr(keys, "=") - 1
            'ListKey1.AddItem (Left(keys, pos))
            GetValues.Add (Mid(keys, pos + 2))
            'ListKey2.AddItem (Mid(keys, pos + 2))
          End If
        Else
          finished = True
        End If
      Loop Until finished = True
    End If
  Loop Until finished = True
  Close File

End Function

Public Function GetAllKeys(mIniFileName As String, ByVal Section As String) As Collection
  
Dim Value As String, retval As String, X As Integer
Dim S() As String, i As Integer
 
  retval = String$(255, 0)
  X = getprivateprofilestring(Section, vbNullString, "", retval, Len(retval), mIniFileName)
  Value = Trim(Left(retval, X))
  S = Split(Value, Chr(0))
  Set GetAllKeys = New Collection
  With GetAllKeys
    For i = LBound(S) To UBound(S)
      If S(i) <> "" Then .Add S(i)
    Next
  End With

End Function

Public Sub SetValue(INIPath As String, INIGroup As String, _
 INIKey As String, INIValue As String)
 
' Update single line of .INI file
' NOTE: group and/or parameter line will be added if not already present
' Input: inipath=path to .INI file, inigroup=group header (no[])
' inikey=parameter name, inivalue=new parameter value
Dim X As Long

  X = writeprivateprofilestring(INIGroup, INIKey, INIValue, INIPath)
 
End Sub

Public Function GetValue(INIPath As String, INIGroup As String, _
 INIKey As String, INIDefault As String) As String
 
' Get line from .INI file
' Input: inpath=path to .INI file, INIgroup=group header (no[])
' Inikey=parameter name, inidefault=default value
' output: value from file or default if not found
Dim X As Long
Dim strBuff As String * 512

  X = getprivateprofilestring(INIGroup, INIKey, INIDefault, strBuff, Len(strBuff), INIPath)
  GetValue = Left$(strBuff, X)
 
End Function




