VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XploraClassic - KramWell.com"
   ClientHeight    =   4785
   ClientLeft      =   1125
   ClientTop       =   1470
   ClientWidth     =   10215
   Icon            =   "app.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4785
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdCommand 
         Caption         =   "CMD"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "PING"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optLocation 
         Caption         =   "C$"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   450
         Width           =   495
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "Default"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   200
         Width           =   855
      End
      Begin VB.OptionButton optLocation 
         Caption         =   "Win 7+"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   200
         Width           =   855
      End
      Begin VB.OptionButton optLocation 
         Caption         =   "Win XP"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "open"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   4320
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Network Path"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton gotoPath 
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbNetlist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "app.frx":08CA
         Left            =   120
         List            =   "app.frx":08CC
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local Drives"
      Height          =   735
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.DriveListBox drvImage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtImage 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   9255
   End
   Begin VB.FileListBox filImage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   4920
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      ReadOnly        =   0   'False
      TabIndex        =   1
      Top             =   960
      Width           =   5175
   End
   Begin VB.DirListBox dirImage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4695
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp1"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuPopUp2 
      Caption         =   "PopUp2"
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh1 
         Caption         =   "&Refresh"
      End
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by KramWell - 22/NOV/2015
'XploraClassic is essentially that. A classic looking windows explorer packed with useful features!

Option Explicit

Private Declare Function IsUserAdmin Lib "shell32" Alias "#680" () As Boolean

Private Const Splitter As String = "<::>"

Dim Data As String
Dim DataPart() As String

Private glPid     As Long
Dim FSys As New FileSystemObject

'this carries over the varible of the folder location
Dim strFpath As String

'Dim FSys As New FileSystemObject
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
'Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Private Sub error_Cancel()
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "Error " & Err.Number
End If
End Sub

Private Sub cmdCommand_Click()

    glPid = Shell("cmd.exe", vbNormalFocus)
        If glPid = 0 Then
            MsgBox "Could not start process", vbExclamation, "Error"
        End If

End Sub

Private Sub cmdOpen_Click()

Dim sTemp As String
Dim isEXE As String

On Error GoTo errorrepair

' Start the process.
sTemp = Trim$(txtImage.Text)
If sTemp = "" Then Exit Sub

'opening folders
If filImage.FileName = "" Then

sTemp = "explorer.exe " & sTemp
        glPid = Shell(sTemp, vbNormalFocus)

    Else

'calculates the ending tag eg go.exe = "exe"
isEXE = filImage.FileName
isEXE = StrReverse$(isEXE) 'reverse filename go.exe > exe.og
isEXE = Mid$(isEXE, 1, 3) 'takes first 3 chars exe.og > exe

'if exe then no need to let explorer open it...
If isEXE = "exe" Then
    glPid = Shell(sTemp, vbNormalFocus)
        Else
    sTemp = "explorer.exe " & sTemp
    glPid = Shell(sTemp, vbNormalFocus)
End If


If glPid = 0 Then
    MsgBox "Could not start process", vbExclamation, "Error"
End If

    End If

'this does 3 things, if filimage.filename isnt selected then it assumes its a folder, if it has text then it
'reverses the filename and sees if that file is an exe or not, if so opens without explorer

errorrepair:
error_Cancel

End Sub

Private Sub cmdPing_Click()
Dim strPingRequest As String

If cmbNetlist.Text <> "" Then

strPingRequest = getHostName(cmbNetlist.Text)

strPingRequest = "cmd.exe /k ping " & strPingRequest & " -t"
    glPid = Shell(strPingRequest, vbNormalFocus)
        If glPid = 0 Then
            MsgBox "Could not start process", vbExclamation, "Error"
        End If

End If
        
End Sub

Private Function getHostName(netValue As String) As String

Dim strString As String

strString = Mid$(netValue, 1, 2) 'takes first 2 chars \\server > \\
If strString = "\\" Then
netValue = Mid(netValue, 3) 'remove first two chars
End If

'split text by \
DataPart = Split(netValue, "\")

getHostName = DataPart(0) 'datapart now holds the value of the servername

End Function

Private Function addSlash(netValue As String) As String
Dim hasBackSlashs As String
'calculates the beginging two chars
hasBackSlashs = Mid$(netValue, 1, 2) 'takes first 2 chars \\server > \\
If hasBackSlashs <> "\\" Then
netValue = "\\" & netValue 'add backslashes
End If

addSlash = netValue

End Function




Private Sub gotoPath_Click()
Dim netValue As String

On Error GoTo errorrepair

If cmbNetlist.Text <> "" Then

netValue = addSlash(cmbNetlist.Text)

If optDefault.Value = True Then
dirImage.Path = netValue
ElseIf optLocation(0).Value = True Then
dirImage.Path = netValue & "\c$"
ElseIf optLocation(1).Value = True Then
dirImage.Path = netValue & "\c$\Users\Public\Desktop"
ElseIf optLocation(2).Value = True Then
dirImage.Path = netValue & "\c$\Documents and Settings\All Users\Desktop"
End If

End If

errorrepair:
error_Cancel
End Sub

Private Sub loadNetworkList()

Dim strNetlist As String
Dim sFileText As String
Dim iFileNo As Integer

strNetlist = App.Path + "\nl.xc1"

'if file netlist exists then read into combo-box
If FSys.FileExists(strNetlist) = True Then
iFileNo = FreeFile
Open strNetlist For Input As #iFileNo
Do While Not EOF(iFileNo)
  Input #iFileNo, sFileText
  cmbNetlist.AddItem sFileText
Loop
Close #iFileNo
End If 'File no exist
End Sub

Private Sub dirImage_Change()
filImage.Path = dirImage.Path
txtImage.Text = "" + filImage.Path
End Sub
Private Sub drvImage_Change()
On Error GoTo errorrepair
dirImage.Path = drvImage.Drive
cmbNetlist.Text = ""
errorrepair:
error_Cancel
End Sub

Private Sub filImage_Click()
Dim strSubstr1 As String

strSubstr1 = Right$(filImage.Path, 1)

If strSubstr1 = "\" Then
txtImage.Text = "" + filImage.Path + filImage.FileName + ""
    Else
txtImage.Text = "" + filImage.Path + "\" + filImage.FileName + ""
    End If

End Sub

Private Sub filImage_DblClick()
Call cmdOpen_Click
End Sub

Private Sub filImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strListSel As String

If Button = vbRightButton Then

    If filImage.ListIndex = -1 Then 'nothing selected
PopupMenu mnuPopUp2, vbPopupMenuRightButton
        'MsgBox "Nothing selected."
            Else

PopupMenu mnuPopUp, vbPopupMenuRightButton

    End If
    End If

End Sub


Private Sub Form_Load()

    Dim rc As Long

    If App.PrevInstance Then
        'rc = MsgBox("Application is already running", vbCritical, App.Title)
        AppActivate "XploraClassic - mwells.info"
        Unload Me
        Exit Sub
    Else
        frmImage.Show
        Call loadNetworkList
        Call loadUsername
        
        dirImage.Path = "c:\"
        txtImage.Text = filImage.Path
        optDefault.Value = True
    End If


End Sub

Private Sub loadUsername()
Dim UserName As String
UserName = Environ("USERNAME")

If IsUserAdmin() = 0 Then
 ' frmImage.Caption = "(NotAdmin-" & UserName & ") " & frmImage.Caption
Else
 ' frmImage.Caption = "(Admin-" & UserName & ") " & frmImage.Caption
End If

End Sub

Private Sub mnuSave_Click()

Dim Destination As String
Dim sPath As String
Dim intListX As Integer
Dim strSubstr1 As String
Dim strSource123 As String
Dim intOK As Integer
Dim answer As Integer
   
   sPath = FixPath(strFpath)
   strFpath = BrowseForFolderByPIDL(sPath)
    
   If strFpath = "" Then
   'dont start loop
        Else
    
    For intListX = filImage.ListCount - 1 To 0 Step -1

        If filImage.Selected(intListX) Then

'this gives the current selected list
'MsgBox filImage.List(intListX)

             If Len(strFpath) = "3" Then
       Destination = strFpath & filImage.List(intListX)
                Else
       Destination = strFpath & "\" & filImage.List(intListX)
                End If

strSubstr1 = Right$(filImage.Path, 1)
If strSubstr1 = "\" Then
strSource123 = "" + filImage.Path + filImage.List(intListX) + ""
    Else
strSource123 = "" + filImage.Path + "\" + filImage.List(intListX) + ""
    End If

'MsgBox Destination & " " & strSource123
'FSys.CopyFile strSource123, Destination

    If FSys.FileExists(Destination) Then
    'do you want to overwrite file
    
    answer = MsgBox("Do you want to overwrite file " & filImage.List(intListX) & "?", vbYesNo)
        If answer = vbYes Then
        'copy file
    FSys.CopyFile strSource123, Destination
        'check to see if copied ok
        
            If FSys.FileExists(Destination) Then
                'copied ok!
                intOK = 1
                Else
                'copy failed
                MsgBox "Failed to copy file: " & filImage.List(intListX)
            End If
        
        End If
    
    Else
    FSys.CopyFile strSource123, Destination
    
    If FSys.FileExists(Destination) Then
    'copied ok
    intOK = 1
    Else
    'Copy failed
    MsgBox "Failed to copy file: " & filImage.List(intListX)
    End If
    
    End If
    
    

        End If
Next
        End If
        
If intOK = 1 Then
    MsgBox "File(s) Saved!"
End If
    
End Sub

Private Sub mnuDelete_Click()
Dim intListX As Integer
Dim strSubstr1 As String
Dim strSource123 As String
Dim Aa As String
Dim intOK As Integer

Aa = 9
intOK = 0

    For intListX = filImage.ListCount - 1 To 0 Step -1

        If filImage.Selected(intListX) Then

strSubstr1 = Right$(filImage.Path, 1)
If strSubstr1 = "\" Then
strSource123 = "" + filImage.Path + filImage.List(intListX) + ""
    Else
strSource123 = "" + filImage.Path + "\" + filImage.List(intListX) + ""
    End If

If Aa = 9 Then
Aa = MsgBox("Are you sure?", vbYesNo)
End If

If Aa = vbNo Then
    Exit Sub
End If

If Aa = vbYes Then

        If FSys.FileExists(strSource123) Then
            FSys.DeleteFile (strSource123)
                If FSys.FileExists(strSource123) Then
                'error deleting file skip, retry or end
                MsgBox "Failed to delete file: " & filImage.List(intListX)
                Else
                'toggle ok,
                'MsgBox "File(s) Deleted"
                intOK = 1
                End If
            
        Else

        MsgBox "File not found: " & filImage.List(intListX)
        End If
End If

        End If
Next
                filImage.Refresh
                dirImage.Refresh
                
If intOK = 1 Then
    MsgBox "File(s) Deleted"
End If

End Sub

Private Sub mnuRefresh1_Click()
'refreshes both filImage & dirImage
        filImage.Refresh
        dirImage.Refresh
End Sub

Private Sub mnuRefresh_Click()
'refreshes both filImage & dirImage
        filImage.Refresh
        dirImage.Refresh
End Sub

Private Sub filImage_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemCount As Integer
Dim lngpos As Integer
Dim i As Integer
Dim Destination As String
Dim intgoTO As Integer
Dim strTest As String
Dim strTest1 As String
 
On Error GoTo errorrepair
 
'On Error Resume Next
ItemCount = Data.Files.Count

intgoTO = 0

 'loop
For i = 1 To ItemCount

'this gets the last charactors after the \
strTest = StrReverse$(Data.Files(i))
lngpos = InStr(strTest, "\")
lngpos = lngpos - 1
strTest = Mid$(strTest, 1, lngpos)
'finding the first(last) dot is what we ideally need
lngpos = InStr(strTest, ".")
strTest = StrReverse$(strTest)

    'if 0 then folder
If lngpos = 0 Then

    If filImage.Path = "" Then
        Else
Destination = filImage.Path & "\" & strTest

If Data.Files(i) = Destination Then
intgoTO = 2
Else
FSys.CopyFolder Data.Files(i), Destination
intgoTO = 1
End If
        End If
    Else
        If filImage.Path = "" Then
            Else
Destination = filImage.Path & "\" & strTest
'copy file

If Data.Files(i) = Destination Then
intgoTO = 2
Else
FSys.CopyFile Data.Files(i), Destination
intgoTO = 1
    End If
        End If
    End If

Next
  
 If intgoTO = 0 Then
  MsgBox "Copy file/folder(s) failed!"
    ElseIf intgoTO = 1 Then
        filImage.Refresh
        dirImage.Refresh
            MsgBox "Folder/File(s) copied OK!"
            End If
'If Err Then Err.Clear

errorrepair:
error_Cancel

End Sub

Private Function BrowseForFolderByPIDL(sSelPath As String) As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim sPath As String * MAX_PATH
     
   With BI
      .hOwner = Me.hwnd
      .pidlRoot = 0
      .lpszTitle = "Select the destination."
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      .lParam = GetPIDLFromPath(sSelPath)
   End With
  
   pidl = SHBrowseForFolder(BI)
  
   If pidl Then
      If SHGetPathFromIDList(pidl, sPath) Then
         BrowseForFolderByPIDL = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      Else
         BrowseForFolderByPIDL = ""
      End If
     
     'free the pidl from SHBrowseForFolder call
      Call CoTaskMemFree(pidl)
   Else
      BrowseForFolderByPIDL = ""
   End If
  
 'free the pidl (lparam) from GetPIDLFromPath call
   Call CoTaskMemFree(BI.lParam)
  
End Function

Private Function GetPIDLFromPath(sPath As String) As Long
        
      GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(sPath, vbUnicode))

End Function

Private Function IsValidDrive(sPath As String) As Boolean

   Dim buff As String
   Dim nBuffsize As Long

   nBuffsize = GetLogicalDriveStrings(0&, buff)

   buff = Space$(nBuffsize)
   nBuffsize = Len(buff)

   If GetLogicalDriveStrings(nBuffsize, buff) Then
   
      IsValidDrive = InStr(1, buff, sPath, vbTextCompare) > 0
   
   End If

End Function


Private Function FixPath(sPath As String) As String

   If Len(sPath) = 0 Then
      FixPath = "C:\"
      Exit Function
   End If
   
   If IsValidDrive(sPath) Then
      FixPath = QualifyPath(sPath)
   Else
      FixPath = UnqualifyPath(sPath)
   End If
   
End Function


Private Function QualifyPath(sPath As String) As String
 
   If Len(sPath) > 0 Then
 
      If Right$(sPath, 1) <> "\" Then
         QualifyPath = sPath & "\"
      Else
         QualifyPath = sPath
      End If
      
   Else
      QualifyPath = ""
   End If
   
End Function


Private Function UnqualifyPath(sPath As String) As String

   If Len(sPath) > 0 Then
   
      If Right$(sPath, 1) = "\" Then
      
         UnqualifyPath = Left$(sPath, Len(sPath) - 1)
         Exit Function
      
      End If
   
   End If
   
   UnqualifyPath = sPath
   
End Function


Private Sub txtImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtImage.ToolTipText = txtImage.Text
End Sub
