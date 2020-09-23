VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add files"
   ClientHeight    =   4830
   ClientLeft      =   1275
   ClientTop       =   1800
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4830
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSpecify 
      Caption         =   "&Specify..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "&Include path from:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtVerification 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin ComctlLib.ProgressBar prgProgress 
      Height          =   135
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search..."
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtSelected 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4695
   End
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDummy 
      Height          =   255
      Left            =   5760
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblVerify 
      Caption         =   "&Verify:"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblAdd 
      Caption         =   "&Files to add:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLevel 
      Caption         =   "C&ompression level:"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkInclude_Click()
  cmdSpecify.Enabled = chkInclude.Value = 1
  txtPath.Enabled = cmdSpecify.Enabled
End Sub

Private Sub cmdAdd_Click()
  Dim i As Long, ip As Long, Num As Long
  Dim ItsOnList As Boolean, ProgOk As Boolean

  If txtPassword.Text <> txtVerification.Text Then
    MsgBox "Password verification does not match", vbCritical
    Exit Sub
  End If
  If lstFiles.ListCount > 100 Then
    prgProgress.Visible = True
    ProgOk = True
     For i = 1 To lstFiles.ListCount
      Num = Num + i + i - 1
    Next i
  End If
  For i = 1 To lstFiles.ListCount
    If Len(txtPath.Text) And LCase(Left(lstFiles.List(i - 1), Len(txtPath.Text))) = LCase(txtPath.Text) And Len(Mid(lstFiles.List(i - 1), Len(txtPath.Text) + 1)) > 255 Then
      If ProgOk Then prgProgress.Visible = False
      lstFiles.ListIndex = i - 1
      MsgBox "File name exceeds 255 characters long", vbCritical
      Exit Sub
    End If
    If FileVerify(lstFiles.List(i - 1)) = False Then
      If ProgOk Then prgProgress.Visible = False
      MsgBox "File '" + lstFiles.List(i - 1) + "' does not exists"
      Exit Sub
    End If
    If LCase(lstFiles.List(i - 1)) = LCase(MainFile) Then
      If ProgOk Then prgProgress.Visible = False
      lstFiles.ListIndex = i - 1
      MsgBox "Adding the archive to the same archive is not allowed" + vbCrLf + "Cannot continue", vbCritical
      Exit Sub
    End If
    For ip = i + 1 To lstFiles.ListCount
      If ProgOk Then Calculate Num, (lstFiles.ListCount - 1) * (i - 1) + ip - i
      If LCase(lstFiles.List(ip - 1)) = LCase(lstFiles.List(i - 1)) Then
        ItsOnList = True
        Exit For
      End If
    Next ip
  Next i
  If ProgOk Then prgProgress.Visible = False
  If ItsOnList Then If MsgBox("Some files are repeated in list" + vbCrLf + "Are you sure you want to continue?", vbYesNo) = vbNo Then Exit Sub
  ReDim Files(lstFiles.ListCount)
  MainPassword = txtPassword.Text
  Select Case cmbLevel.ListIndex
    Case 0
      i = 0
    Case 1
      i = 3
    Case 2
      i = 4
    Case 3
      i = 6
    Case 4
      i = 8
    Case 5
      i = 9
  End Select
  If chkInclude.Value = 1 Then
    MainPath = txtPath.Text
    If Right(MainPath, 1) <> "\" Then MainPath = MainPath + "\"
  Else
    MainPath = ""
  End If
  Files(0) = Format(i)
  For i = 0 To lstFiles.ListCount - 1
    Files(i + 1) = lstFiles.List(i)
  Next i
  Unload Me

End Sub

Private Sub cmdBrowse_Click()
  Dim File As String, TmpStr() As String
  Dim i As Integer, ip As Integer
  Dim ItsOnList As Boolean

  File = Dialog(Me, 2)
  If Len(File) = 0 Then Exit Sub
  TmpStr = Split(File, Chr(0))
  If UBound(TmpStr) <> 0 Then
    If Right(TmpStr(0), 1) <> "\" Then TmpStr(0) = TmpStr(0) + "\"
    For i = 1 To UBound(TmpStr)
      TmpStr(i) = TmpStr(0) + TmpStr(i)
      ItsOnList = False
      For ip = 1 To lstFiles.ListCount
        If LCase(lstFiles.List(ip - 1)) = LCase(TmpStr(i)) Then
          ItsOnList = True
          Exit For
        End If
      Next ip
      If ItsOnList = False Then lstFiles.AddItem TmpStr(i)
    Next i
  Else
    For ip = 1 To lstFiles.ListCount
      If LCase(lstFiles.List(ip - 1)) = LCase(TmpStr(0)) Then
        ItsOnList = True
        Exit For
      End If
    Next ip
    If ItsOnList = False Then lstFiles.AddItem TmpStr(0)
  End If
  If lstFiles.ListCount > 0 Then cmdAdd.Enabled = True
  cmdRemove.Enabled = False
End Sub

Private Sub cmdCancel_Click()
  ReDim Files(0)
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  WinHelp Me.hwnd, HelpFile, &H1, 20
End Sub

Private Sub cmdRemove_Click()
  lstFiles.RemoveItem lstFiles.ListIndex
  txtSelected.Text = ""
  cmdRemove.Enabled = False
  If lstFiles.ListCount = 0 Then cmdAdd.Enabled = False
End Sub

Private Sub cmdSearch_Click()
  frmSearch.Show vbModal
End Sub

Private Sub cmdSpecify_Click()
  frmSpecify.Show vbModal
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim Folder As String

'  cmbLevel.ListIndex = 2
  If UBound(Files) <> 0 Then
    For i = 1 To UBound(Files)
      If Form1.ExistDir(Files(i)) Then
        Search Files(i), "*.*", lstFiles, lblDummy
        If Len(Folder) = 0 Then Folder = Files(i)
      ElseIf FileVerify(Files(i)) Then
        lstFiles.AddItem Files(i)
      End If
    Next i
    If lstFiles.ListCount <> 0 Then cmdAdd.Enabled = True
  End If
  If Len(Folder) Then
    chkInclude.Value = 1
    txtPath.Text = Folder
  End If
End Sub

Private Sub lstFiles_Click()
  txtSelected.Text = lstFiles.List(lstFiles.ListIndex)
  txtSelected.SelStart = Len(txtSelected.Text)
  cmdRemove.Enabled = True
End Sub

Private Sub txtSelected_Change()
  Dim TmpInt As Integer

  TmpInt = lstFiles.ListIndex
  If TmpInt < 0 Then Exit Sub
  lstFiles.RemoveItem TmpInt
  lstFiles.AddItem txtSelected.Text, TmpInt
  lstFiles.ListIndex = TmpInt
End Sub

Sub Calculate(Val1 As Long, Val2 As Long)
  Dim TmpInt As Integer

  TmpInt = Int(Val2 / Val1 * 100)
  If Val1 < Val2 Or prgProgress.Value = TmpInt Then Exit Sub
  prgProgress.Value = TmpInt
  DoEvents
End Sub

Private Sub txtSelected_KeyPress(KeyAscii As Integer)
  If lstFiles.ListIndex = -1 And KeyAscii = 13 Then
    KeyAscii = 0
    lstFiles.AddItem txtSelected.Text
    txtSelected.Text = ""
  End If
End Sub
