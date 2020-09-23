VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileDB"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   1710
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraAddFiles 
      Height          =   4560
      Left            =   180
      TabIndex        =   5
      Top             =   360
      Width           =   5280
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   465
         Left            =   2745
         TabIndex        =   16
         Top             =   3645
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   465
         Left            =   1170
         TabIndex        =   15
         Top             =   3645
         Width           =   1455
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   330
         Left            =   4410
         TabIndex        =   14
         Top             =   405
         Width           =   690
      End
      Begin VB.PictureBox picFlat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   1575
         ScaleHeight     =   585
         ScaleWidth      =   1665
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtFileDescription 
         Height          =   330
         Left            =   135
         MaxLength       =   255
         TabIndex        =   9
         Top             =   855
         Width           =   4200
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the file to add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label lblFile 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   450
         Width           =   4200
      End
   End
   Begin VB.Frame fraViewFiles 
      Height          =   4560
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   5280
      Begin VB.Frame fraViewFile 
         Height          =   870
         Left            =   3825
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   630
         Width           =   1185
      End
      Begin VB.Frame fraDelete 
         Height          =   870
         Left            =   3825
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   2340
         Width           =   1185
      End
      Begin MSComctlLib.ListView lvwDB 
         Height          =   4020
         Left            =   6435
         TabIndex        =   2
         Top             =   2160
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   7091
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tvwDB 
         Height          =   4245
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   7488
         _Version        =   393217
         Indentation     =   12
         Style           =   7
         ImageList       =   "imlIcons"
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag a file here to delete it"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   3825
         TabIndex        =   10
         Top             =   1890
         Width           =   1410
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMain 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag a file here to view it"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   3825
         TabIndex        =   6
         Top             =   225
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMain 
         Caption         =   "File.mdb contents"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6255
         TabIndex        =   4
         Top             =   3150
         Width           =   2535
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4590
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "File.mdb contents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2835
      TabIndex        =   0
      Top             =   270
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open File"
      End
      Begin VB.Menu mnuSaveToDatabase 
         Caption         =   "&Save to database"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuAddFiles 
         Caption         =   "&Add Files"
      End
      Begin VB.Menu mnuViewFiles 
         Caption         =   "&View Files"
      End
      Begin VB.Menu mnuPurge 
         Caption         =   "&Purge Files"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout"
      Begin VB.Menu mnuShowAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As clsConnection
Dim mSupportedFileTypes As String
Dim mDragNode As Node, mHilitNode As Node
Dim mContinue As Boolean


Private Sub cmdBrowse_Click()
    ShowOpenDialog
End Sub

Private Sub cmdCancel_Click()
    lblFile.Caption = ""
    txtFileDescription.Text = ""
    Me.Refresh
End Sub

Private Sub cmdSave_Click()
    If txtFileDescription.Text <> "" Then
        SaveFileToTable (lblFile.Caption)
    Else
        Call MsgBox("Please enter a description for this file.", vbInformation, App.EXEName)
        txtFileDescription.SetFocus
    End If
End Sub



Private Sub Form_Load()
    If App.PrevInstance Then End
   ' Set con = New clsConnection
    If Not InitConnection Then
    Else
        ToggleMenu (False)
       ' InitControls
        InitFrames
        LoadImageList
        If mContinue = False Then End
        initTvwDB

        LoadTvwDB
        LoadSupportedFiletypes
        Call MakeFlatButtons(Me)
        
    End If
End Sub
Private Function InitConnection() As Boolean
    Set con = New clsConnection
    If Not con.openConnection Then
        MsgBox "Could not open database"
        Set con = Nothing
        End
    Else
        InitConnection = True
    End If
End Function
Private Sub ToggleMenu(b As Boolean)
    mnuOpenFile.Enabled = b
    mnuSaveToDatabase.Enabled = b
End Sub
Private Sub InitControls()


End Sub
Private Sub InitFrames()
    fraViewFiles.Visible = True
    fraAddFiles.Visible = False
    fraAddFiles.Width = fraViewFiles.Width
    fraAddFiles.Height = fraViewFiles.Height
    fraViewFiles.Left = 0
    fraViewFiles.Top = 0
    fraAddFiles.Top = 0
    fraAddFiles.Left = 0
    
End Sub
Private Sub LoadImageList()

Dim rs As New ADODB.Recordset
Dim i As Integer
Dim imgX As ListImage
Dim recordcount As Integer
On Error GoTo TheError
    mContinue = True
    imlIcons.ListImages.Clear
'    imlIcons.ImageHeight = 32
'    imlIcons.ImageWidth = 32
    rs.Open "Select filetype.* from filetype where status=1", con.pConn, adOpenStatic, adLockOptimistic
    recordcount = rs.recordcount
    Do Until i = recordcount
        i = i + 1
        Set imgX = imlIcons.ListImages.Add(i, rs.Fields("filetypedesc"), LoadPicture(App.Path & "\icon\" & rs.Fields("fileiconpath")))
        rs.MoveNext
    Loop
    Set imgX = imlIcons.ListImages.Add(i + 1, "All", LoadPicture(App.Path & "\icon\All Files.ico"))
Cleanup:
    rs.Close
    Set rs = Nothing
    Exit Sub
TheError:
    If err.Number = 76 Then
        Call MsgBox("Could not find the icon for " & rs.Fields("filetypedesc") & ".", vbInformation, App.Title)
        mContinue = False
        Resume Cleanup
    End If
End Sub
Private Sub initTvwDB()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim nodeX As Node
Dim k As String
Dim icon As String
Dim recordcount As Integer
    With tvwDB
        .ImageList = imlIcons
        .Nodes.Clear
        .LineStyle = tvwRootLines
        Set nodeX = .Nodes.Add(, , "r", "All Files", "All")
        rs.Open "Select filetype.* from filetype where status=1", con.pConn, adOpenStatic, adLockOptimistic
        recordcount = rs.recordcount
        Do Until i = recordcount
            i = i + 1
            k = "C" & CStr(rs.Fields("filetypeId"))
            icon = rs.Fields("Filetypedesc")
            Set nodeX = .Nodes.Add("r", tvwChild, k, icon, icon)
            rs.MoveNext
        Loop
        rs.Close
    Set rs = Nothing
    End With
End Sub

Private Sub LoadTvwDB()
Dim rs As New ADODB.Recordset
Dim key As String
Dim nodeX As Node
Dim k As String
Dim i As Integer
    rs.Open "SELECT File.fileID,File.FileType, File.FileDesc, FileType.FileTypeDesc FROM File INNER JOIN FileType ON File.FileType = FileType.FiletypeID where file.status=1;", con.pConn
    Do Until rs.EOF
        i = i + 1
        key = rs.Fields("filetypedesc")
        k = "C" & rs.Fields("filetype")
        Set nodeX = tvwDB.Nodes.Add(k, tvwChild, "F" & rs.Fields("Fileid"), rs.Fields("FileDesc"), key)
        rs.MoveNext
    Loop
    tvwDB.Nodes(1).Expanded = True
    'nodeX.Expanded = True
    rs.Close
    Set rs = Nothing
End Sub

Private Sub LoadSupportedFiletypes()
Dim rs As New ADODB.Recordset
    mSupportedFileTypes = "Supported File types|"
    rs.Open "Select filetypedesc,filetypeext from filetype where status=1", con.pConn
    Do Until rs.EOF
        mSupportedFileTypes = mSupportedFileTypes & "*." & rs.Fields("filetypeext") & ";"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim iAns As Integer
iAns = MsgBox("Are you sure you want to exit?", vbQuestion + _
    vbYesNo, "Exit?")
    
    If iAns = vbNo Then Cancel = True
End Sub





Private Sub fraDelete_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim key As Long
Dim ret As String
Dim rs As New ADODB.Recordset
    If mDragNode Is Nothing Then
        Call MsgBox("Please select a file to drag.", vbInformation, App.EXEName)
    Else
        key = CLng(Replace(mDragNode.key, "F", ""))
        ret = MsgBox("Do you want to delete " & mDragNode.Text & "?", vbYesNo, App.EXEName)
        If ret = vbYes Then
            rs.Open "Select status from file where status=1 and fileid=" & key, con.pConn, adOpenKeyset, adLockOptimistic
            rs.Fields("Status") = 8
            rs.Update
            rs.Close
            Set mDragNode = Nothing
            initTvwDB
            LoadTvwDB
        End If
    End If
    Set rs = Nothing
End Sub








Private Sub fraViewFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim fileID As Long
Dim rs As New ADODB.Recordset
Dim FileStream As New ADODB.Stream
Dim ext As String
Dim fileName As String
    If mDragNode Is Nothing Then
        Call MsgBox("Please select a file to view.", vbInformation, App.EXEName)
    Else
        Screen.MousePointer = vbHourglass
        fraViewFile.Enabled = False
        fileName = App.Path & "\" & App.EXEName & "temp."
        fileID = Mid(mDragNode.key, 2)
        rs.Open "Select file.*,filetype.filetypeext from file inner join filetype on file.filetype=filetype.filetypeid where fileID=" & fileID, con.pConn
        If Not rs.EOF Then
            ext = rs.Fields("filetypeext")
            fileName = fileName & ext
            FileStream.Type = adTypeBinary
            FileStream.Open
            FileStream.Write rs.Fields("Filedata").Value
            FileStream.SaveToFile fileName, adSaveCreateOverWrite
        End If
        rs.Close
        fraViewFile.Enabled = True
        Screen.MousePointer = vbDefault
        Call LaunchAppFile(Me, fileName)
    End If
    
    Set rs = Nothing
End Sub







Private Sub mnuAddFiles_Click()
    fraViewFiles.Visible = False
    fraAddFiles.Visible = True
    lblFile.Caption = ""
    txtFileDescription.Text = ""
    ToggleMenu (True)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpenFile_Click()
    ShowOpenDialog
End Sub

Private Sub mnuPurge_Click()
Dim ret As Long

    ret = MsgBox("Do you want to purge the database?", vbYesNo, App.EXEName)
    If ret = vbYes Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        con.pConn.Execute "delete file.* from file where status=8"
        con.pConn.Close
        Set con = Nothing
        On Error Resume Next
        Kill App.Path & "\" & App.EXEName & "temp.*"
        
        If CompactDatabase(App.Path & "\file.mdb") Then
           Call MsgBox("Database purged.", vbInformation, App.EXEName)
        Else
            Call MsgBox("Could not compact the database.", vbExclamation, App.EXEName)
        End If
        If Not InitConnection Then
            Call MsgBox("Cannot reopen database.", vbExclamation, App.EXEName)
        End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuShowAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuViewFiles_Click()
    fraViewFiles.Visible = True
    fraAddFiles.Visible = False
    ToggleMenu (False)
End Sub
Function CompactDatabase(strFileName As String) As Boolean
    Dim objJro As JRO.JetEngine
    Dim objFileSystem As FileSystemObject
    Dim strTmpFileName As String
   On Error GoTo EXIT_PROC
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    strTmpFileName = objFileSystem.GetSpecialFolder(TemporaryFolder).Path & "\" & objFileSystem.GetFileName(strFileName)
    Set objJro = New JRO.JetEngine
    
    objJro.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFileName & ";jet OLEDB:Database password=Scope4", _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTmpFileName & ";Jet OLEDB:Engine Type=5"
    
    objFileSystem.CopyFile strTmpFileName, strFileName
    objFileSystem.DeleteFile strTmpFileName, True
    
    CompactDatabase = True
    Exit Function
EXIT_PROC:
    
    Set objFileSystem = Nothing
    Set objJro = Nothing
End Function





Private Sub tvwDB_DragOver(Source As Control, x As Single, y As Single, State As Integer)

    If Not mDragNode Is Nothing Then
        tvwDB.DropHighlight = tvwDB.HitTest(x, y)
    End If

End Sub

Private Sub tvwDB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Set mDragNode = tvwDB.HitTest(x, y)
    If Mid(mDragNode.key, 1, 1) <> "F" Then
        Set mDragNode = Nothing
    End If
End Sub

Private Sub tvwDB_NodeClick(ByVal Node As MSComctlLib.Node)
    Exit Sub
    If Mid(Node.key, 1, 1) = "F" Then MsgBox Node.key
End Sub

Private Sub ShowOpenDialog()
On Error GoTo TheErr
    lblFile.Caption = ""
    txtFileDescription.Text = ""
    Cdlg.Filter = mSupportedFileTypes
    Cdlg.ShowOpen
    lblFile.Caption = Cdlg.fileName
    txtFileDescription.SetFocus
    Exit Sub
TheErr:
    lblFile.Caption = ""
    txtFileDescription.Text = ""
End Sub
Private Sub ToggleLVButtons()
    cmdSave.Enabled = Not cmdSave.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled
End Sub
Private Sub SaveFileToTable(f As String)
Dim rs As New ADODB.Recordset
Dim fileTypeRs As New ADODB.Recordset
Dim pos As Integer
Dim fileTypeID As Long
Dim fileTypeStr As String
Dim fileID As Long
Dim FileStream As New ADODB.Stream
On Error GoTo TheError
    ToggleLVButtons
    pos = InStr(1, f, ".")
    If pos > 1 Then
        fileTypeStr = UCase(Mid(f, pos + 1))
        
        fileTypeRs.Open "Select filetypeid from filetype where ucase(filetypeext)='" & fileTypeStr & "'", con.pConn
        fileTypeID = fileTypeRs.Fields("FiletypeID")
        
        fileTypeRs.Close
    End If
    'check to see if a file exists
    rs.Open "Select fileLocation from file where filelocation='" & f & "' and status=1", con.pConn
    If rs.EOF Then
        rs.Close
    
        rs.Open "select file.* from file where status=-1", con.pConn, adOpenKeyset, adLockOptimistic
        rs.AddNew
        rs.Fields("Filedesc").Value = txtFileDescription.Text
        rs.Fields("filetype") = fileTypeID
        rs.Fields("fileLocation") = f
        rs.Fields("Status").Value = 1
        rs.Update
        rs.Close
        rs.Open "Select max(fileID)as maxID from file where status=1", con.pConn
        fileID = rs.Fields("maxid")
        rs.Close
        FileStream.Type = adTypeBinary
        FileStream.Open
        rs.Open "Select file.* from file where fileid=" & fileID, con.pConn, adOpenKeyset, adLockOptimistic
        FileStream.LoadFromFile f

        rs.Fields("Filedata") = FileStream.Read
        rs.Update
        rs.Close
        initTvwDB
        LoadTvwDB
        Call MsgBox(f & " was saved successfully.", vbInformation, App.EXEName)
    Else
        rs.Close
        Call MsgBox(f & " is already present in the database", vbInformation, App.EXEName)
    End If
Finish:

    
    Set rs = Nothing
    Set fileTypeRs = Nothing
    Set FileStream = Nothing
    ToggleLVButtons
    
    Exit Sub
TheError:
    If rs.State = adStateOpen Then rs.Close
    If fileTypeRs.State = adStateOpen Then fileTypeRs.Close
    If err.Number = 3002 Then
        con.pConn.Execute "delete file.* from file where fileid=" & fileID
        Call MsgBox("Could not read that file.  Try again.", vbExclamation, App.EXEName)
    Else
        MsgBox err.Description
    End If
    
    GoTo Finish
End Sub

Private Sub tvwDB_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error Resume Next
    If mDragNode.Parent Is Nothing Then Set mDragNode = Nothing
End Sub
