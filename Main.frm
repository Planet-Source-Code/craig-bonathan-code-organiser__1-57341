VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Organiser"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ProjectDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select a Visual Basic Project"
      Filter          =   "Visual Basic Project (*.vbp)|*.vbp"
   End
   Begin VB.CommandButton OrganiseButton 
      Caption         =   "Organise"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   5415
   End
   Begin VB.CommandButton GenerateButton 
      Caption         =   "Open Project"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.ListBox FileList 
      Height          =   2535
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FSO As Object
Dim FileCollection As Collection

Private Sub Form_Load()
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub GenerateButton_Click()
    Dim Pos As Long, FileNum As Long, CurrentLine As String, LineType As String, FileName As String
    Dim FileNameSplit() As String ', Temp As Long
    
    On Error GoTo ExitOnCancel::
    ProjectDialog.ShowOpen
    FileName = ProjectDialog.FileName
    On Error GoTo 0
    
    FileList.Clear
    Set FileCollection = Nothing
    Set FileCollection = New Collection
    
    FileNum = FreeFile
    Open FileName For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, CurrentLine
        Pos = InStr(1, CurrentLine, "=")
        If Pos > 1 And Len(CurrentLine) > 0 Then
            LineType = UCase(Mid(CurrentLine, 1, Pos - 1))
            CurrentLine = Mid(CurrentLine, Pos + 1)
            FileNameSplit = Split(CurrentLine, ";")
            CurrentLine = Trim(FileNameSplit(UBound(FileNameSplit)))
            If FSO.GetDriveName(CurrentLine) = "" Then
                CurrentLine = FSO.BuildPath(FSO.GetParentFolderName(FileName), CurrentLine)
            End If
            CurrentLine = FSO.GetAbsolutePathName(CurrentLine)
            Select Case LineType
            Case "FORM"
                FileList.AddItem "Form: " & CurrentLine
            Case "MODULE"
                FileList.AddItem "Module: " & CurrentLine
            Case "CLASS"
                FileList.AddItem "Class: " & CurrentLine
            End Select
        End If
    Loop
    Close #FileNum
    
    For Pos = 0 To FileList.ListCount - 1
        FileList.List(Pos) = Format(Pos + 1, "000") & " " & FileList.List(Pos)
        FileNameSplit = Split(FileList.List(Pos), ":", 2)
        FileName = Trim(FileNameSplit(1))
        FileCollection.Add FileName, CStr(Pos)
        FileNameSplit = Split(FileList.List(Pos), ":")
        FileList.List(Pos) = FileNameSplit(0) & ": " & FSO.GetBaseName(FileName)
        FileList.Selected(Pos) = True
    Next
    
    FileList.ListIndex = -1
    
ExitOnCancel::
End Sub

Private Sub OrganiseButton_Click()
    Dim Pos As Long, FileName As String, SuccessCount As Long, TotalCount As Long
    
    FileList.Enabled = False
    
    For Pos = 0 To FileList.ListCount - 1
        FileList.ListIndex = Pos
        If FileList.Selected(Pos) = True Then
            TotalCount = TotalCount + 1
            FileName = FileCollection.Item(CStr(FileList.ListIndex))
RetryFile::
            If FSO.FileExists(FileName) = False Then
                Select Case MsgBox("""" & FileName & """ does not exist", vbCritical Or vbAbortRetryIgnore)
                Case vbAbort
                    Exit For
                Case vbRetry
                    GoTo RetryFile::
                Case vbIgnore
                    GoTo SkipToNext::
                End Select
            End If
            ScanFile FileName
        End If
        SuccessCount = SuccessCount + 1
SkipToNext::
    Next
    
    FileList.ListIndex = -1
    
    FileList.Enabled = True
End Sub

Private Sub ScanFile(FileName As String)
    Dim FileNum As Long, CurrentTab As Long, LineData As String, TempAntiTab As Long, TempTab As Long
    Dim TabArray() As Long, LineNum As Long, LineCode As String, LineComment As String
    Dim OutputString As String
    
    FileNum = FreeFile
    Open FileName For Input As #FileNum
    Do While Not EOF(FileNum)
        ReDim TabArray(LineNum)
        Line Input #FileNum, LineData
        
        SeparateLine LineData, LineCode, LineComment
        RemoveSpaces LineCode
        RemoveSpaces LineComment
        If LineComment <> "" And LineCode <> "" Then LineComment = " " & LineComment
        
        TempAntiTab = 0
        
        If BeginsWith(LineCode, "End Sub") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "End Function") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "Next") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "Loop") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "Wend ") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "End If") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "End With") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "End Select") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "End Type") = True Then CurrentTab = CurrentTab - 1
        If BeginsWith(LineCode, "End Enum") = True Then CurrentTab = CurrentTab - 1
        
        If BeginsWith(LineCode, "Case ") = True Then TempAntiTab = 1
        If BeginsWith(LineCode, "ElseIf ") = True Then TempAntiTab = 1
        If LineCode = "Else" Then TempAntiTab = 1
        
        If CurrentTab < 0 Then CurrentTab = 0
        OutputString = OutputString & Space((CurrentTab - TempAntiTab + TempTab) * 4) & LineCode & LineComment & vbCrLf
        
        TempTab = 0
        
        If BeginsWith(LineCode, "Public Sub") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Private Sub") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Public Function") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Private Function") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Function ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Sub ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "For ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Do ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "While ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Select Case ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "With ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Public Type ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Private Type ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Type ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Public Enum ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Private Enum ") = True Then CurrentTab = CurrentTab + 1
        If BeginsWith(LineCode, "Enum ") = True Then CurrentTab = CurrentTab + 1
        
        If LineCode <> "" Then
            If Mid(LineCode, Len(LineCode), 1) = "_" Then TempTab = 2
        End If
        If Len(LineCode) > 5 Then
            If BeginsWith(LineCode, "If ") = True And _
                    Mid(LineCode, Len(LineCode) - 4, 5) = " Then" Then CurrentTab = CurrentTab + 1
        End If
    Loop
    Close #FileNum
    
    If Len(OutputString) > 2 Then
        OutputString = Mid(OutputString, 1, Len(OutputString) - 2)
    End If
    
    FileNum = FreeFile
    Open FileName For Output As #FileNum
    Print #FileNum, OutputString
    Close #FileNum
End Sub

Private Sub SeparateLine(Text As String, Code As String, Comment As String)
    Dim Pos As Long, Quotes As Boolean, CommentMode As Boolean, Letter As String * 1
    Code = ""
    Comment = ""
    For Pos = 1 To Len(Text)
        Letter = Mid(Text, Pos, 1)
        If CommentMode = False Then
            If Letter = """" Then
                Quotes = Not Quotes
                Code = Code & Letter
            ElseIf Letter = "'" And Quotes = True Then
                Code = Code & Letter
            ElseIf Letter = "'" And Quotes = False Then
                CommentMode = True
                Comment = Comment & Letter
                Else
                Code = Code & Letter
            End If
            Else
            Comment = Comment & Letter
        End If
    Next
End Sub

Private Sub RemoveSpaces(Text As String)
    Dim Pos As Long, Characters As Boolean
    
    For Pos = 1 To Len(Text)
        If Mid(Text, Pos, 1) <> " " Then Characters = True
    Next
    If Characters = False Then
        Text = ""
        Exit Sub
    End If
    Do While Mid(Text, 1, 1) = " "
        Text = Mid(Text, 2)
    Loop
    For Pos = Len(Text) To 1 Step -1
        If Mid(Text, Pos, 1) = " " Then
            Text = Mid(Text, 1, Pos - 1)
            Else
            Exit For
        End If
    Next
End Sub

Function BeginsWith(Text As String, Beginning As String) As Boolean
    If Len(Text) < Len(Beginning) Then Exit Function
    If UCase(Mid(Text, 1, Len(Beginning))) = UCase(Beginning) Then BeginsWith = True
End Function
