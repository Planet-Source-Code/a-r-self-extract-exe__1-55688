VERSION 5.00
Begin VB.Form frmPack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Files"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmPack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk 
      Caption         =   "Delete exe after the install"
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set as autorun"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove File"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New File"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private autorun As String

Dim ItemLength As Long
Dim ItemString As String
Dim ItemNumber(0 To 1) As Integer

Dim BytesExtract As String
Dim BytesAdd As String

Dim ItemBinary As String
Dim Position As Long
Dim LastPosition As Long

Dim FileListStart As Long
Dim FilePosition As Long
Dim ExitDo As Boolean

Dim PutLength As String
Dim PutPosition As Long

Function JpkAdd(JpkFile As String, filename As String, AddName As String) As Boolean

    On Error GoTo FinaliseError

    AddName = AddName & Chr(0)
    
    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As #ItemNumber(0)
        ItemNumber(1) = FreeFile
        Open filename For Binary As #ItemNumber(1)
            PutLength = LOF(ItemNumber(1)) & Chr(0)
            Put ItemNumber(0), LOF(ItemNumber(0)) + 1, AddName
            Put ItemNumber(0), LOF(ItemNumber(0)) + 1, PutLength
            PutPosition = LOF(ItemNumber(0))
            If LOF(ItemNumber(1)) > 1000000 Then
                Position = -999999
                Do
                    Position = Position + 1000000
                    If Position + 999999 > LOF(ItemNumber(1)) Then BytesAdd = String(LOF(ItemNumber(1)) - Position + 1, Chr$(0)) Else BytesAdd = String(1000000, Chr$(0))
                    Get ItemNumber(1), Position, BytesAdd
                    Put ItemNumber(0), PutPosition + 1, BytesAdd
                    PutPosition = LOF(ItemNumber(0))
                Loop Until Position + 999999 > LOF(ItemNumber(1))
            Else
                BytesAdd = String(LOF(ItemNumber(1)), Chr$(0))
                Get ItemNumber(1), , BytesAdd
                Put ItemNumber(0), PutPosition + 1, BytesAdd
            End If
        Close ItemNumber(1)
    Close #ItemNumber(0)
    JpkAdd = True
    Exit Function
    
FinaliseError:
    JpkAdd = False

End Function

Function JpkList(JpkFile As String, ListItem As ListBox) As Boolean

    On Error GoTo FinaliseError

    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As #ItemNumber(0)
        Position = 1
        Do
            ItemString = Space(256)
            Get #ItemNumber(0), Position, ItemString
            ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
            Position = Position + Len(ItemString) + 1
            ListItem.AddItem ItemString
            
            ItemString = Space(256)
            Get #ItemNumber(0), Position, ItemString
            ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
            ItemLength = CLng(ItemString)
            Position = Position + Len(ItemString) + ItemLength + 1
        Loop Until Position > LOF(ItemNumber(0))
    Close #ItemNumber(0)
    JpkList = True
    Exit Function
    
FinaliseError:
    JpkList = False

End Function

Function JpkExtract(JpkFile As String, filename As String, Destination As String) As Boolean

    On Error GoTo FinaliseError

    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As ItemNumber(0)
        ItemNumber(1) = FreeFile
        Open Destination For Binary As ItemNumber(1)
            Position = 1
            ExitDo = False
            Do
                ItemString = Space(256)
                Get #ItemNumber(0), Position, ItemString
                ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
                Position = Position + Len(ItemString) + 1
                If LCase(ItemString) = LCase(filename) Then ExitDo = True
                
                ItemString = Space(256)
                Get #ItemNumber(0), Position, ItemString
                ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
                ItemLength = CLng(ItemString)
                Position = Position + Len(ItemString) + ItemLength + 1
                If ExitDo = True Then Exit Do
            Loop Until Position > LOF(ItemNumber(0))
            
            FileListStart = Position - ItemLength
            If ItemLength > 1000000 Then
                FilePosition = -999999
                Do
                    FilePosition = FilePosition + 1000000
                    If FilePosition + 999999 > ItemLength Then BytesExtract = Space(ItemLength - FilePosition + 1) Else BytesExtract = Space(1000000)
                    Get ItemNumber(0), FileListStart, BytesExtract
                    Put ItemNumber(1), FilePosition, BytesExtract
                    FileListStart = FileListStart + Len(BytesExtract)
                Loop Until FilePosition + 999999 > LOF(ItemNumber(1))
            Else
                BytesExtract = Space(ItemLength)
                Get ItemNumber(0), Position - ItemLength, BytesExtract
                Put ItemNumber(1), 1, BytesExtract
            End If
        Close ItemNumber(1)
    Close ItemNumber(0)
    JpkExtract = True
    Exit Function
    
FinaliseError:
    JpkExtract = False

End Function

Private Sub Command1_Click()
 Dim filename As String
    filename = modFolder.ShowOpenDialog("All Files|*.*", 0, "Add File", "", "")
    If filename = "" Then Exit Sub
    List1.AddItem filename
End Sub

Private Sub Command2_Click()
    If List1.ListIndex <> -1 Then If MsgBox("Are you sure you want to remove the selected file?", vbQuestion + vbYesNo, "Question") = vbYes Then List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
 Dim i As Long
 Dim ff As Integer
 Dim filename As String
 
    Screen.MousePointer = 11
    For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        JpkAdd App.Path & "\tmpPack.pkz", List1.Text, GetFileName(List1.Text)
    Next
    ff = FreeFile
    
    Open App.Path & "\autorun.ini" For Output As #ff
        Print #ff, GetFileName(autorun)
        Print #ff, CBool(chk.Value)
        DoEvents
    Close #ff
    
    JpkAdd App.Path & "\tmpPack.pkz", App.Path & "\autorun.ini", "autorun.ini"
    Kill App.Path & "\autorun.ini"
    
    Screen.MousePointer = 1
    DoEvents
    
    filename = modFolder.ShowSaveDialog("Exe Files|*.exe", 0, "Save Exe", "", "")
    DoEvents
    
    If filename = "" Then Exit Sub
    Screen.MousePointer = 11
    If Right(filename, 4) <> ".exe" Then filename = filename & ".exe"
    
    Call NewExe(filename, LoadData(App.Path & "\tmpPack.pkz"))
    Kill App.Path & "\tmpPack.pkz"
    Screen.MousePointer = 1
    MsgBox "Done"
End Sub

Private Sub Command4_Click()
    If List1.ListIndex <> -1 Then If MsgBox("Are you sure you want the selected file will be the autorun?", vbQuestion + vbYesNo, "Question") = vbYes Then autorun = List1.Text
End Sub

Private Sub Form_Load()
    ReDim Files(0)
End Sub

Private Function GetFileName(Path As String) As String
Dim Findsep As Long

    For Findsep = 1 To Len(Path)
        If Mid(Path, Len(Path) - (Findsep - 1), 1) = "\" Or Mid(Path, Len(Path) - (Findsep - 1), 1) = "/" Then
            GetFileName = Right(Path, Findsep - 1)
            Exit Function
        End If
    Next Findsep

End Function

Private Function GetPath(FullPath As String) As String
    
    Dim c As Integer
    Dim s As Integer
    Dim j As Integer
    Dim m As Long
    Dim GetChr0 As String, GetChr1 As String
    
    c = 0: s = 0: j = 0
    
    For m = 1 To Len(FullPath)
        GetChr0 = Right(FullPath, m): GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then c = c + 1
    Next m
    For m = 1 To Len(FullPath)
        GetChr0 = Left(FullPath, m): GetChr1 = Right(GetChr0, 1)
        j = j + 1
        If GetChr1 = "\" Or GetChr1 = "/" Then
            j = 0: s = s + 1
            If s = c Then GetPath = Right(GetChr0, m - j): Exit Function
        End If
    Next m

End Function

