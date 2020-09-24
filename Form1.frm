VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INIEdit"
   ClientHeight    =   6468
   ClientLeft      =   120
   ClientTop       =   576
   ClientWidth     =   7716
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6468
   ScaleWidth      =   7716
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Add New"
      Height          =   2292
      Left            =   2640
      TabIndex        =   9
      Top             =   1800
      Width           =   5052
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   372
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   852
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   960
         TabIndex        =   16
         Top             =   1320
         Width           =   3852
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   960
         TabIndex        =   14
         Top             =   840
         Width           =   3852
      End
      Begin VB.CommandButton Command3 
         Caption         =   "New"
         Height          =   252
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   732
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   3012
      End
      Begin VB.Label Label4 
         Caption         =   "Value:"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "Section:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   972
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2292
      Left            =   2640
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   5052
      _ExtentX        =   8911
      _ExtentY        =   4043
      _Version        =   393216
      Cols            =   3
      AllowUserResizing=   1
   End
   Begin VB.Frame frameValue 
      Caption         =   "Value"
      Height          =   1335
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   5052
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   1320
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   360
         Visible         =   0   'False
         Width           =   168
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Edit"
         Height          =   372
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   852
      End
      Begin VB.TextBox txtValue 
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Value:"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4812
      End
      Begin VB.Label cptValueName 
         Caption         =   "Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4812
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   3480
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      Filter          =   "INI Files (*.ini) | *.ini"
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3360
      Top             =   2280
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":300A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3166
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   264
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7716
      _ExtentX        =   13610
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New INI File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open INI File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save INI File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newval"
            Object.ToolTipText     =   "New Value"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   3960
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3856
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6012
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   10605
      _Version        =   393217
      Indentation     =   176
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewDoc 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuspec1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuspec2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Item"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuLaunchNotepad 
         Caption         =   "&Launch Notepad"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuShowGrid 
         Caption         =   "&Show Grid"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSortWindow 
         Caption         =   "S&ort"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelpMain 
      Caption         =   "&Help"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuABouyt 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim ctitle As String
Dim curfile As String
Dim cNode As String
Dim cParent As String
Dim fsagent As New FileSystemObject
Dim dox As Boolean
Const HKEY_CLASSES_ROOT = &H80000000

Public Sub RefreshINI(strFPath As String)
Dim Y As Integer: Y = 1
Do Until Y > TreeView1.Nodes.Count
     TreeView1.Nodes.Remove (Y)
     Y = Y + 1
Loop
Y = 0
Do Until Combo1.ListCount = 0
     Combo1.RemoveItem (0)
     Combo1.Refresh
Loop
Do Until Y > grid.Rows - 1
     grid.Row = Y
     grid.Col = 1
     grid.Text = ""
     grid.Col = 2
     grid.Text = ""
     Y = Y + 1
Loop
grid.Rows = 1
ctitle = Space(999)
Dim ret As Long
ret = GetFileTitle(strFPath, ctitle, 999)
curfile = strFPath
TreeView1.Nodes.Add , , "INITOP", ctitle, 3, 3

Dim strlines(9999) As String
Dim trash As Integer
'load file into array
     trash = -1
     Open strFPath For Input As #1
          Do Until EOF(1)
               Line Input #1, strlines(trash + 1)
               trash = trash + 1
          Loop
     Close #1
'check file for values
Dim cur, x As Integer
Dim curhost As String
Dim str As String
Dim tchar As String
cur = 0
     Do Until cur > trash
          str = strlines(cur)
          str = Trim(str)
          If Left(str, 1) = "[" Then
               x = 0
               tchar = ""
               Do Until Right(tchar, 1) = "]"
                    x = x + 1
                    tchar = tchar & Mid(str, x, 1)
               Loop
               tchar = Left(tchar, Len(tchar) - 1)
               tchar = Right(tchar, Len(tchar) - 1)
               TreeView1.Nodes.Add "INITOP", tvwChild, "[" & tchar, tchar, 1, 1
               Combo1.AddItem tchar
               curhost = tchar
          ElseIf Left(str, 1) = ";" Then
     
          ElseIf str = "" Then
          
          Else
               x = 0
               tchar = ""
               Do Until Right(tchar, 1) = "="
                    If x >= Len(str) Then
                         'throw out
                         Exit Sub
                    End If
                    x = x + 1
                    tchar = tchar & Mid(str, x, 1)
               Loop
               tchar = Left(tchar, Len(tchar) - 1)
               TreeView1.Nodes.Add "[" & curhost, tvwChild, "*" & CStr(grid.Rows), tchar, 2, 2
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Rows - 1
                grid.Col = 0
                grid.Text = curhost
                grid.Col = 1
                grid.Text = tchar
               Dim retval As Long, xt As String
               xt = Space(256)
               retval = GetPrivateProfileString(curhost, tchar, "", xt, 256, curfile)
               grid.Col = 2
               grid.Text = CStr(Trim(xt))
          End If
          cur = cur + 1
     Loop
TreeView1.Nodes(1).Expanded = True
End Sub

Private Sub Command1_Click()
Load frmNewValue
frmNewValue.Text1.Text = txtValue.Text
Me.Enabled = False
frmNewValue.Show
End Sub

Private Sub Command2_Click()
DeleteKey (TreeView1.SelectedItem.key)
End Sub

Private Sub Command3_Click()
Dim nv As New frmNewValue
nv.Caption = "New Section"
nv.Label1.Caption = "Name:"
Load nv
nv.Show
End Sub

Private Sub Command4_Click()
Dim retval As Long
If Trim(Combo1.Text) = "" Then
     Label2.ForeColor = vbRed
     Exit Sub
Else
     Label2.ForeColor = vbBlack
     If Trim(Text1.Text) = "" Then
          Label3.ForeColor = vbRed
          Exit Sub
     Else
          Label3.ForeColor = vbBlack
          retval = WritePrivateProfileString(Combo1.Text, Text1.Text, Text2.Text, curfile)
          DoEvents
          RefreshINI (curfile)
          Text1.Text = ""
          Text2.Text = ""
     End If
End If
TreeView1.Sorted = True

End Sub

Private Sub Form_Load()
TreeView1.Sorted = True
If RegFileAssociation(".ini", App.Path & "\" & App.EXEName, True) = False Then
     MsgBox "Unable to register filetype. INIEdit is not associated with the windows shell.", vbExclamation + vbOKOnly + vbApplicationModal + vbDefaultButton1, "Shell Error"
End If
If Trim(Command()) <> "" Then
     RefreshINI (Trim(Command()))
End If
grid.Cols = 3
grid.ColWidth(0) = Len("Section     ") * 100
grid.Col = 0
grid.Row = 0
grid.Text = "Section"
grid.Col = 1
grid.ColWidth(1) = Len("Name        ") * 100
grid.Text = "Name"
grid.Col = 2
grid.ColWidth(2) = Len("Value        ") * 5000
grid.Text = "Value"
End Sub

Private Sub mnuABouyt_Click()
MsgBox "INI Edit v1.1" & vbCrLf & _
vbCrLf & "Primary Development by Will Urbanski" & vbCrLf _
 & "(C) 2001 Mango Vision LLC  [ http://www.mangovision.com ]" & vbCrLf & vbCrLf _
 & "For more information visit our website or mail [ will@mangovision.com ]", vbInformation + vbOKOnly + vbApplicationModal + vbDefaultButton1, "About"
End Sub

Private Sub mnuAction_Click()
If fsagent.FileExists(curfile) = False Then
     Me.mnuLaunchNotepad.Enabled = False
     Me.mnuRefresh.Enabled = False
     Me.mnuShowGrid.Enabled = False
     Me.mnuDelete.Enabled = False
Else
     Me.mnuLaunchNotepad.Enabled = True
     Me.mnuRefresh.Enabled = True
     Me.mnuShowGrid.Enabled = True
     Me.mnuDelete.Enabled = True
End If
End Sub

Private Sub mnuClose_Click()
Dim Y As Integer: Y = 1
Do Until Y > TreeView1.Nodes.Count
     TreeView1.Nodes.Remove (Y)
     Y = Y + 1
Loop
Y = 0
Do Until Combo1.ListCount = 0
     Combo1.RemoveItem (0)
     Combo1.Refresh
Loop
txtValue.Text = ""
cptValueName.Caption = "Name: "
Y = 0
Do Until Y > grid.Rows - 1
     grid.Row = Y
     grid.Col = 1
     grid.Text = ""
     grid.Col = 2
     grid.Text = ""
     Y = Y + 1
Loop
grid.Rows = 1
End Sub

Private Sub mnuDelete_Click()
DeleteKey (TreeView1.SelectedItem.key)
End Sub

Private Sub mnuExit_Click()
Unload Form1
End
End Sub

Private Sub mnuLaunchNotepad_Click()
If fsagent.FileExists(curfile) = False Then Exit Sub
Shell "Notepad.exe " & curfile, vbNormalFocus
End Sub

Private Sub mnuNewDoc_Click()
     Call mnuClose_Click
     With CommonDialog1
          .FileName = ""
          .ShowSave
          If .FileName = "" Then Exit Sub
          Open .FileName For Output As #1
          Print #1, ";; " & .FileName & " ;;"
          Close #1
          Form1.Caption = "INIEdit - [" & .FileName & "]"
          RefreshINI (.FileName)
     End With
End Sub

Private Sub mnuOpen_Click()
     Call mnuClose_Click
     With CommonDialog1
          .FileName = ""
          .ShowOpen
          If .FileName = "" Then Exit Sub
          'i/o stuff
          If fsagent.FileExists(.FileName) = False Then
               MsgBox "Error 75 raised, the file " & .FileName & " was unable to be opened.", vbCritical + vbOKOnly + vbSystemModal + vbDefaultButton1, "75"
          Else
               RefreshINI (.FileName)
               Form1.Caption = "INIEdit - [ " & .FileName & " ]"
          End If
     End With
End Sub

Private Sub mnuRefresh_Click()
If fsagent.FileExists(curfile) = False Then Exit Sub
RefreshINI (curfile)
End Sub

Private Sub mnuSave_Click()
'it saves automatically
End Sub

Private Sub mnuSaveAs_Click()
If fsagent.FileExists(curfile) = False Then Exit Sub
With Me.CommonDialog1
     .FileName = ""
     .ShowSave
     If .FileName = "" Then Exit Sub
     FileCopy curfile, .FileName
End With
End Sub

Private Sub mnuShowGrid_Click()
If mnuShowGrid.Checked = False Then
     grid.Visible = True
     mnuShowGrid.Checked = True
Else
     grid.Visible = False
     mnuShowGrid.Checked = False
End If
End Sub

Private Sub mnuSortWindow_Click()
If mnuSortWindow.Checked = True Then
     TreeView1.Sorted = False
     mnuSortWindow.Checked = False
Else
     TreeView1.Sorted = True
     mnuSortWindow.Checked = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
     Case "new"
          mnuNewDoc_Click
     Case "open"
          mnuOpen_Click
     Case "save"
          mnuSave_Click
     Case "delete"
          If fsagent.FileExists(curfile) = False Then Exit Sub
          DeleteKey (TreeView1.SelectedItem.key)
     Case "newval"
          Call Command3_Click
     Case "refresh"
          If fsagent.FileExists(curfile) = False Then Exit Sub
          RefreshINI (curfile)
End Select
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim retval As Long
If Left(Node.key, 1) = "[" Then
     'throw out
Else
     Dim gh As String
     gh = Space(999)
     Dim preforce As String
     Dim x As Integer, tchar As String
     If Node.key = "INITOP" Then Exit Sub
     Dim str As Integer
     tchar = Node.key
     tchar = Right(tchar, Len(tchar) - 1)
     str = CInt(tchar)
     grid.Row = str
     grid.Col = 0
     preforce = grid.Text
     grid.Col = 1
     tchar = grid.Text
     cNode = CStr(str)
     Text3.Text = CStr(cNode)
     cParent = CStr(preforce)
     retval = GetPrivateProfileString(preforce, tchar, "", gh, 999, curfile)
     cptValueName.Caption = "Name: " & tchar
     txtValue.Text = CStr(gh)
End If
End Sub
Public Sub DeleteKey(key As String)
     If fsagent.FileExists(curfile) = False Then Exit Sub
     If key = "INITOP" Then Exit Sub
     Dim preforce As String
     Dim x As Integer, tchar As String
     Dim str As Integer
     If Left(Trim(key), 1) = "[" Then
          'its a header, delete everything
          Dim trsh(9999) As String
          Dim indx, ci As Integer
          Dim nva As Boolean: nva = False
          Dim msgb
          msgb = MsgBox("Are you sure you want to delete the section '" & Right(key, Len(key) - 1) & "'?", vbQuestion + vbYesNo + vbDefaultButton2 + vbApplicationModal, "Delete")
          Select Case msgb
          Case vbYes
               Open curfile For Input As #1
                    indx = indx + 1
                    Do Until EOF(1)
                         Line Input #1, trsh(indx)
                              If Left(Trim(trsh(indx)), 1) = "[" Then
                                   nva = False
                              End If
                              If nva = True And Left(Trim(trsh(indx)), 1) <> "[" Then
                                   trsh(indx) = ""
                              End If
                              If Trim(trsh(indx)) = key & "]" Then
                                   'start to destroy
                                   trsh(indx) = ""
                                   nva = True
                              End If
                         indx = indx + 1
                    Loop
               Close #1
          
          ci = 1
          Open curfile For Output As #1
               Do Until ci = indx
                    Print #1, trsh(ci)
                    ci = ci + 1
               Loop
          Close #1
          Case vbNo
               Exit Sub
          End Select
     Else
          tchar = key
     tchar = Right(tchar, Len(tchar) - 1)
     str = CInt(tchar)
     grid.Row = CInt(str)
     grid.Col = 1
     preforce = grid.Text
     grid.Col = 2
     tchar = grid.Text
     cNode = CStr(str)
     Text3.Text = CStr(cNode)
     cParent = CStr(preforce)
     Dim cx As String
     cx = TreeView1.SelectedItem.Parent.key
          'we have all the info needed
          Dim trash(9999) As String
          Dim ind As Integer
          Dim cnt, nv As Boolean
          cnt = False
          nv = False
          Open curfile For Input As #1
               ind = ind + 1
               Do Until EOF(1)
                    Line Input #1, trash(ind)
                    If nv = False Then
                         If cnt = False Then
                              If Trim(trash(ind)) = "[" & Trim(preforce) & "]" Then
                                   cnt = True
                              End If
                         ElseIf cnt = True Then
                              If Left(trash(ind), Len(tchar)) = tchar Then
                                   trash(ind) = ""
                                   nv = True
                              End If
                         End If
                    End If
               ind = ind + 1
               Loop
          Close #1
     Open curfile For Output As #1
     str = 1
     Do Until str = ind
          Print #1, trash(str)
          str = str + 1
     Loop
     Close #1
     End If
     RefreshINI (curfile)
End Sub
