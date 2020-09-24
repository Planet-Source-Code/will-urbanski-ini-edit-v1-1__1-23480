VERSION 5.00
Begin VB.Form frmNewValue 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Value"
   ClientHeight    =   1164
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewValue.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1164
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "&Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmNewValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub Command1_Click()
If Left(Me.Caption, 4) = "Edit" Then
     Form1.txtValue.Text = Text1.Text
     Dim retval As Long
     Form1.grid.Col = 2
     Form1.grid.Row = CInt(Form1.Text3.Text)
     retval = WritePrivateProfileString(cParent, Form1.grid.Text, Form1.txtValue.Text, curfile)
Else
     If Trim(Text1.Text) = "" Then Unload Me
     Form1.Combo1.AddItem Text1.Text
     Me.Caption = "Edit Value"
     Me.Label1.Caption = "Value:"
     Form1.Combo1.ListIndex = Form1.Combo1.ListCount - 1
     Form1.Combo1.Refresh
End If
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
Me.Hide
End Sub
