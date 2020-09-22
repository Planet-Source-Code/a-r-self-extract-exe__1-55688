VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picProgbar 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.Label lblPrgbar 
         BackColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while the setup unpacks the files..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Sub SetBar(precent)
    DoEvents
    lblPrgbar.Width = picProgbar.Width / 100 * precent
    lblPrgbar.Visible = IIf(precent = 0, False, True)
    DoEvents
End Sub

Private Sub UnloadApp(ByRef delete As Boolean)
 Dim ff As Integer
 
    If Not (delete) Then End: Exit Sub
    ff = FreeFile
    
    Open App.Path & "\tmp.bat" For Output As #ff
        Print #ff, "@echo off"
        Print #ff, "del " & App.EXEName & ".exe"
        Print #ff, "del tmp.bat"
    Close #ff
    
    Shell App.Path & "\tmp.bat", vbHide
    End
End Sub

Private Sub Form_Load()
Dim i As Long
Dim ff As Integer
Dim auto As String, str As String
Dim deleteme As Boolean

    SetBar 0
    Me.Show

    If Not (SaveData(App.Path & "\tmp~006.pkz")) Then
        MsgBox "The program can not run whiout any file in it.", vbCritical, "Error"
        End
    End If
    
    SetBar 40

    JpkList App.Path & "\tmp~006.pkz", List1
    SetBar 50
    
    For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        JpkExtract App.Path & "\tmp~006.pkz", List1.Text, App.Path & "\" & List1.Text
        SetBar 50 + ((20 / List1.ListCount) * (i + 1))
    Next
    
    ff = FreeFile
    
    Open App.Path & "\autorun.ini" For Input As #ff
        Input #ff, auto
        Input #ff, str
        deleteme = CBool(str)
    Close #ff
    
    SetBar 90
    
    Kill App.Path & "\autorun.ini"
    Kill App.Path & "\tmp~006.pkz"
    
    SetBar 100

    Call modShellExec.ShellExecLaunchFile(auto, App.Path & "\", "")
    Call UnloadApp(deleteme)
End Sub

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

Function JpkExtract(JpkFile As String, FileName As String, Destination As String) As Boolean

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
                If LCase(ItemString) = LCase(FileName) Then ExitDo = True
                
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
