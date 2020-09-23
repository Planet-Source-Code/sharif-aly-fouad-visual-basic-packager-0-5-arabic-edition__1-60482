VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ãÔÑæÚ  ÇáÓíÊÈ ÇáÐÇÊí _ V.B.4.Arab_Auto_Setup _without Setup GUI ÈÏæä æÇÌåå ÇáÓíÊÈ"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8865
   Icon            =   "vbsetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   5040
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "EXE Files(exe) |*.exe"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbsetup.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbsetup.frx":111E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path"
         Object.Width           =   9225
      EndProperty
   End
   Begin ComctlLib.ListView lstFiles 
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2355
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1455
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Packed"
         Object.Width           =   1455
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ratio"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modified"
         Object.Width           =   2673
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   3704
      EndProperty
   End
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7435
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   5520
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÇáãßÊÈÇÊ  - OCX's AND DLL's"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   8655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ãáÝÇÊ ÇáÈÑäÇãÌ"
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   8655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"vbsetup.frx":2172
      Height          =   855
      Left            =   5640
      TabIndex        =   8
      Top             =   240
      Width           =   3255
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   4560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   39
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbsetup.frx":2200
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbsetup.frx":3842
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbsetup.frx":4E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbsetup.frx":64C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbsetup.frx":7B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbsetup.frx":914A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning ÊÝÍÕ ÇáãáÝ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   540
      Index           =   0
      Left            =   240
      Picture         =   "vbsetup.frx":A78C
      Top             =   1080
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Image nback 
      Height          =   450
      Index           =   3
      Left            =   120
      Picture         =   "vbsetup.frx":E0B2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image nNext 
      Height          =   450
      Index           =   3
      Left            =   4320
      Picture         =   "vbsetup.frx":10B07
      Top             =   5280
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   4125
   End
   Begin VB.Image nNext 
      Height          =   450
      Index           =   2
      Left            =   1320
      Picture         =   "vbsetup.frx":13584
      Top             =   5040
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image nNext 
      Height          =   450
      Index           =   1
      Left            =   3120
      Picture         =   "vbsetup.frx":163CE
      Top             =   4920
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image nNext 
      Height          =   450
      Index           =   0
      Left            =   7440
      Picture         =   "vbsetup.frx":1921E
      Top             =   4980
      Width           =   1305
   End
   Begin VB.Image nback 
      Height          =   450
      Index           =   2
      Left            =   3720
      Picture         =   "vbsetup.frx":1BC9B
      Top             =   720
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image nback 
      Height          =   450
      Index           =   1
      Left            =   2400
      Picture         =   "vbsetup.frx":1EB30
      Top             =   720
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image nback 
      Height          =   450
      Index           =   0
      Left            =   6165
      Picture         =   "vbsetup.frx":219A5
      Top             =   4980
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   1035
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   6045
   End
   Begin VB.Image Image2 
      Height          =   540
      Index           =   2
      Left            =   6120
      Picture         =   "vbsetup.frx":243FA
      Top             =   0
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Image Image2 
      Height          =   540
      Index           =   1
      Left            =   3480
      Picture         =   "vbsetup.frx":27D20
      Top             =   0
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Menu File 
      Caption         =   "ãáÝ"
      Begin VB.Menu New_pro 
         Caption         =   "ãÔÑæÚ ÌÏíÏ"
      End
      Begin VB.Menu Open_Project 
         Caption         =   "ÝÊÍ ãÔÑæÚ"
         Enabled         =   0   'False
      End
      Begin VB.Menu Exit 
         Caption         =   "ÎÑæÌ"
      End
   End
   Begin VB.Menu Project 
      Caption         =   "ÇáãÔÑæÚ"
      Begin VB.Menu Select_exe 
         Caption         =   "ÇÎÊíÇÑ ãáÝ ÇáãÔÑæÚ _ Exe"
         Enabled         =   0   'False
      End
      Begin VB.Menu Add_files 
         Caption         =   "ÇÖÇÝå ãáÝÇÊ"
         Enabled         =   0   'False
      End
      Begin VB.Menu Add_Lib 
         Caption         =   "ÇÖÇÝå ãßÊÈÇÊ"
         Begin VB.Menu From_EXE 
            Caption         =   "ãä ãáÝ Exe"
            Enabled         =   0   'False
         End
         Begin VB.Menu From_VBPROJ 
            Caption         =   "ãä ãÔÑæÚ ÝíÌæÇá ÈíÓß"
            Enabled         =   0   'False
         End
         Begin VB.Menu files_From_files 
            Caption         =   "ãä ãáÝÇÊ"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu Opti 
         Caption         =   "ÎíÇÑÇÊ"
         Enabled         =   0   'False
      End
      Begin VB.Menu Compile_exe 
         Caption         =   "ÊÍæíá Çáí ãáÝ Exe"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Help 
      Caption         =   "ãÓÇÚÏå"
      Begin VB.Menu About 
         Caption         =   "Íæá ÇáÈÑäÇãÌ"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim LIBFiles() As String, Hands As Integer
Dim nNexts As Integer
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hwnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type
Const AppName As String = "VB4arabPacker"


Public Sub Get_EXE_Data(PathEXE_FileName As String)
Dim nfree As Integer, Exe_Data As String
Dim i As Long, nSum As Long
Hands = 0
nfree = FreeFile
Open PathEXE_FileName For Binary As nfree
     Exe_Data = Space(LOF(nfree))
Get #nfree, , Exe_Data
Close nfree

ProgressBar1.Max = Len(Exe_Data) + 3

For i = 1 To Len(Exe_Data)
    ' DLL..ááÍÕæá Úáì ÃÓãÇÁ ãáÝÇÊ ÇáÜÜ
    If UCase(Mid(Exe_Data, i, 4)) = ".DLL" Then
        For nfree = 1 To 10
           If Asc(Left(Mid(Exe_Data, i - nfree, 4), 1)) = 0 Then
              nfree = nfree - 1
              ReDim Preserve LIBFiles(Hands)
              nSum = (i + 4) - (i - nfree)
              LIBFiles(Hands) = Left(Mid(Exe_Data, i - nfree, i + 4), nSum)
              Hands = Hands + 1
           Exit For
           End If
       Next nfree
    End If
    ' OCX..ááÍÕæá Úáì ÃÓãÇÁ ãáÝÇÊ ÇáÜÜ
    If UCase(Mid(Exe_Data, i, 4)) = ".OCX" Then
       For nfree = 1 To 10
           If Asc(Left(Mid(Exe_Data, i - nfree, 4), 1)) = 0 Then
              nfree = nfree - 1
              ReDim Preserve LIBFiles(Hands)
              nSum = (i + 4) - (i - nfree)
              LIBFiles(Hands) = Left(Mid(Exe_Data, i - nfree, i + 4), nSum)
              Hands = Hands + 1
           Exit For
           End If
       Next nfree
    End If
    
ProgressBar1.Value = i
Next i
End Sub


Public Function CheckDoublValue(nValueChecking As String) As Boolean
CheckDoublValue = False
Dim Eid As Integer

For Eid = 1 To ListView1.ListItems.Count

    If Trim(UCase(ListView1.ListItems.Item(Eid))) = Trim(UCase(nValueChecking)) Then
        CheckDoublValue = True
       Exit For
    End If

Next Eid

End Function


Public Function GetWindowSystemDirectory() As String
    Dim sSave As String, Ret As Long
    sSave = Space(255)
    Ret = GetSystemDirectory(sSave, 255)
    sSave = Left$(sSave, Ret)
    GetWindowSystemDirectory = sSave
End Function


Private Sub About_Click()
frmSplash.Show
End Sub

Private Sub Add_files_Click()
  Dim TmpInt As Integer, i As Integer
  Dim Position As Long, TmpLng As Long, FileLength As Long, CurrentPos As Long
  Dim Buff() As Byte
  Dim TmpStr As String
  Dim Info As proInfo

  If Len(MainFile) = 0 Then
    MsgBox "No open archive", vbCritical
    Exit Sub
  End If
  MainPath = ""
  frmAdd.Show vbModal
  If UBound(Files) = 0 Then Exit Sub
  If FileVerify(MainFile) = False Then
    ReDim Files(0)
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  Open MainFile For Binary Access Read Write As #1
  Get #1, , TmpInt
  If 32767 - UBound(Files) < TmpInt Then
    MsgBox "Archive do not support so many files", vbCritical
    Close
    Exit Sub
  End If
  TmpInt = TmpInt + UBound(Files)
  Seek #1, 1
  Put #1, , TmpInt
  Seek #1, LOF(1) + 1
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Adding files"
  For i = 1 To UBound(Files)
    FileLength = FileLength + FileLen(Files(i))
  Next i
  For i = 1 To UBound(Files)
    Open Files(i) For Binary Access Read As #2
    frmProgress.lblOperation.Caption = "Adding " + Simplificate(Files(i)) + "..."
    DoEvents
    Info.Size = LOF(2)
    Info.Modified = FileDateTime(Files(i))
    Info.Packed = 0
    If Len(MainPath) = 0 Then
      TmpStr = Simplificate(Files(i))
    Else
      If LCase(Left(Files(i), Len(MainPath))) = LCase(MainPath) Then
        TmpStr = Mid(Files(i), Len(MainPath) + 1)
      Else
        TmpStr = Simplificate(Files(i))
      End If
    End If
    ToFile 2, Chr(Len(TmpStr)) + TmpStr
    Position = Seek(1)
    Put #1, , Info
    While Loc(2) < LOF(2)
      ReDim Buff(GetSize(2))
      Get #2, , Buff()
      TmpLng = UBound(Buff)
      CompressArray Buff, Val(Files(0))
      If Len(MainPassword) <> 0 Then
        CodeArray Buff(), StrConv(MainPassword, vbFromprocode)
        ToFile 0, -UBound(Buff)
      Else
        ToFile 0, UBound(Buff)
      End If
      ToFile 0, TmpLng
      Put #1, , Buff()
      Info.Packed = Info.Packed + UBound(Buff) + 9
      frmProgress.Calculate FileLength, CurrentPos + Seek(2)
      DoEvents
    Wend
    CurrentPos = CurrentPos + LOF(2)
    ToFile 0, 0
    Info.Packed = Info.Packed + 4
    Seek #1, Position
    Put #1, , Info
    Seek #1, LOF(1) + 1
    Close #2
  Next i
  Me.Enabled = True
  Unload frmProgress
  Close
  ReDim Files(0)
  LoadFiles MainFile, False
End Sub

Private Sub Addtofinal()
'Label2.Label
ListView1.ListItems.Add 1, , Label1.Caption
For i = 1 To ListView1.ListItems.Count
  lstFiles.ListItems.Add i, , ListView1.ListItems(i).Text
Next
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
nNexts = 0
' New added
CreateDll
  ReDim Files(0)
  HelpFile = App.Path
  If Right(HelpFile, 1) <> "\" Then HelpFile = HelpFile + "\"

'New File

End Sub







Private Sub lstFiles_DblClick()
  Dim i As Integer
  Dim Ret As Long
  Dim Modified As Variant
  Dim TmpStr As String, FileToExtract As String
  Dim FileInfo As SHELLEXECUTEINFO

  If Len(MainFile) = 0 Then Exit Sub
  ExtractPath = Environ("temp")
  If Right(ExtractPath, 1) <> "\" Then ExtractPath = ExtractPath + "\"
  TmpStr = ExtractPath
  mnuExtract_Click
  For i = 1 To lstFiles.ListItems.Count
    FileToExtract = TmpStr + lstFiles.ListItems(i).Text
    If lstFiles.ListItems(i).Selected And FileVerify(FileToExtract) Then
      With FileInfo
        .cbSize = Len(FileInfo)
        .fMask = &H40
        .hwnd = Me.hwnd
        .lpVerb = "open"
        .lpFile = FileToExtract
        .lpParameters = ""
        .lpDirectory = TmpStr
        .nShow = 1
      End With
      Modified = FileDateTime(FileToExtract)
      ShellExecuteEx FileInfo
      WaitForSingleObject FileInfo.hProcess, -1
      If FileVerify(FileToExtract) Then
        If Modified <> FileDateTime(FileToExtract) Then If MsgBox("File has been modified" + vbCrLf + "Update?", vbYesNo + vbQuestion) = vbYes Then UpdateFile i, FileToExtract, lstFiles.ListItems(i).SubItems(5)
        Kill FileToExtract
      End If
      lstFiles.ListItems(i).Selected = True
      Exit Sub
    End If
  Next i
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Image2(0).Picture = Image2(1).Picture
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Image2(0).Picture = Image2(2).Picture
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
   GetEXEFileIcon
   nNext(0).Picture = nNext(1).Picture
   nNexts = 1
Else
   nNext(0).Picture = nNext(0).Picture
   nNexts = 0
End If
End Sub

Private Sub GetEXEFileIcon()
    Dim mIcon As Long
    mIcon = ExtractAssociatedIcon(App.hInstance, CommonDialog1.FileName, 2)
    Picture1.Picture = LoadPicture()
    DrawIconEx Picture1.hdc, 0, 0, mIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon mIcon
    Label1.Caption = CommonDialog1.FileName
End Sub



Private Sub nback_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If nNexts = 2 Then
   nback(0).Picture = nback(2).Picture
End If
End Sub

Private Sub nback_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If nNexts = 2 Then
   nback(0).Picture = nback(2).Picture
   ListView1.Visible = False
   Label2.Visible = False
   nback(0).Picture = nback(3).Picture
   Label1.Caption = ""
   Picture1.Picture = LoadPicture()
   Image2(0).Visible = True
   nNexts = 0
End If

End Sub

Private Sub New_pro_Click()
'ãÔÑæÚ ÌÏíÏ


  Dim File As String

  File = Dialog(Me, 1)
  If Len(File) = 0 Then Exit Sub
  Open File For Output As #1
  Close
  Open File For Binary Access Write As #1
  ToFile 1, 0
  ToFile 1, 0
  Close
  Me.Caption = AppName + " - " + Simplificate(File)
  MainFile = File
  LoadFiles MainFile, False
  
  ' ÌÚá ÇãßÇäíå ÇÎÊíÇÑ ÇáãáÝ ãÊÇÍå
  
    Select_exe.Enabled = True
 ' ÇÊÇÍå ÇáÒÑ ÇÖÇÝå ãáÝÇÊ
 Image2(0).Visible = True
  ' ÇÊÇÍå ÇááÓÊå ÇáËÇäíå
  lstFiles.Visible = True
    
End Sub

Private Sub nNext_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If nNexts = 1 Then nNext(0).Picture = nNext(2).Picture
End Sub

Private Sub nNext_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If nNexts = 1 Then
   Screen.MousePointer = 11

 
 
 
Get_EXE_Data (CommonDialog1.FileName)
   ListView1.ListItems.Clear
   Image2(0).Visible = False
   Picture1.Visible = False
   ProgressBar1.Visible = True
   Label2.Visible = True
   Label2.Caption = Label1.Caption
       ' ãÔÛá ÇáÊÇíãÑ Çáãáæä æ ÇáäÕ
 'Label3.Visible = True
 Timer1.Enabled = True
  ProgressBar1.Visible = True

For i = 0 To Hands - 1
   
    If CheckDoublValue(LIBFiles(i)) = False Then
    
    If UCase(Right(LIBFiles(i), 3)) = "DLL" Then
       ListView1.ListItems.Add 1, "A" & i, LIBFiles(i), 1, 1
       ListView1.ListItems("A" & i).Checked = True
    End If
    If UCase(Right(LIBFiles(i), 3)) = "OCX" Then
       ListView1.ListItems.Add , "A" & i, LIBFiles(i), 2, 2
       ListView1.ListItems("A" & i).Checked = True
    End If
    
    If Dir(GetWindowSystemDirectory & "\" & LIBFiles(i), vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
       ListView1.ListItems("A" & i).SubItems(1) = Format(FileLen(GetWindowSystemDirectory & "\" & LIBFiles(i)), "###,###,###")
       ListView1.ListItems("A" & i).SubItems(2) = GetWindowSystemDirectory
    Else
       ListView1.ListItems("A" & i).SubItems(1) = "0"
       ListView1.ListItems("A" & i).SubItems(2) = "None"
    End If
    
    End If
    
Next i
   

ListView1.Visible = True
nNext(0).Picture = nNext(3).Picture
nback(0).Picture = nback(2).Picture
If ListView1.ListItems.Count > 0 Then
   ListView1.ListItems.Item(1).Selected = True
   ListView1.SetFocus
End If
nNexts = 2
End If
Screen.MousePointer = 0
End Sub




Private Sub Select_exe_Click()
Image2(0).Picture = Image2(2).Picture
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
   GetEXEFileIcon
   nNext(0).Picture = nNext(1).Picture
   nNexts = 1
Else
   nNext(0).Picture = nNext(0).Picture
   nNexts = 0
End If
End Sub

Private Sub Timer1_Timer()
'Label3.ForeColor = &H80000012
'If Label3.ForeColor = &H80000012 Then:
'Timer1.Enabled = False
'Timer2.Enabled = True

If ListView1.Visible = True Then:
Timer1.Enabled = False
ProgressBar1.Visible = False
Label3.Visible = False
Add_files.Enabled = True: Addtofinal

'ÇÖÇÝå ßá ÇáãæÌæÏ ÈäÊÇÆÌ ãä áÓÊ Ýíæ 1
' Çáí
' 1srtfiles
'lstFiles.ListItems.Add = ListView1.ListItems.Item
End Sub

Private Sub Timer2_Timer()
'Label3.ForeColor = &HFF&
'If Label3.ForeColor = &HFF& Then:
'Timer1.Enabled = True
'Timer2.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer

  Select Case KeyCode
    Case vbKeyE
      If Shift = 2 Then
        For i = 1 To lstFiles.ListItems.Count
          lstFiles.ListItems(i).Selected = True
        Next i
        lstFiles_MouseUp 1, 0, 0, 0
      End If
    Case vbKeyI
      If Shift = 2 Then
        For i = 1 To lstFiles.ListItems.Count
          lstFiles.ListItems(i).Selected = Not lstFiles.ListItems(i).Selected
        Next i
        lstFiles_MouseUp 1, 0, 0, 0
      End If
  End Select
End Sub

'Private Sub Form_Resize()
 ' On Error Resume Next
  'lstFiles.Width = Me.ScaleWidth
'  lstFiles.Height = Me.ScaleHeight - lstFiles.Top - stbStatus.Height
'End Sub

Private Sub lstFiles_BeforeLabelEdit(Cancel As Integer)
  Cancel = 1
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim TmpVar As Variant
  Dim TotalSize As Long
  Dim i As Integer

  TmpVar = GetSelected()
  For i = 1 To UBound(TmpVar)
    If TmpVar(i) Then TotalSize = TotalSize + Val(lstFiles.ListItems(i).SubItems(1))
  Next i
  If Val(stbStatus.Panels(2).Text) <> TotalSize Then stbStatus.Panels(2).Text = Format(Int(TotalSize / 1024)) + " KB"
  If Val(stbStatus.Panels(4).Text) <> TmpVar(0) Then stbStatus.Panels(4).Text = Format(TmpVar(0)) + " file(s) selected"
  If Button = 2 Then
    mnuActions_Click
    PopupMenu mnuActions
  End If
End Sub

Private Sub lstFiles_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer

  ReDim Files(Data.Files.Count)
  For i = 1 To UBound(Files)
    Files(i) = Data.Files.Item(i)
  Next i
  mnuAdd_Click
  Me.SetFocus
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuActions_Click()
  Dim TmpVar As Variant

  TmpVar = GetSelected()
  mnuAdd.Enabled = Len(MainFile) <> 0
  mnuComments.Enabled = mnuAdd.Enabled
  mnuTest.Enabled = lstFiles.ListItems.Count <> 0
  mnuExtract.Enabled = mnuAdd.Enabled And TmpVar(0) <> 0
  mnuDelete.Enabled = mnuExtract.Enabled
  mnuViewer.Enabled = mnuDelete.Enabled
  mnuMakeExe.Enabled = mnuAdd.Enabled
End Sub

Private Sub mnuAdd_Click()
  Dim TmpInt As Integer, i As Integer
  Dim Position As Long, TmpLng As Long, FileLength As Long, CurrentPos As Long
  Dim Buff() As Byte
  Dim TmpStr As String
  Dim Info As proInfo

  If Len(MainFile) = 0 Then
    MsgBox "No open archive", vbCritical
    Exit Sub
  End If
  MainPath = ""
  frmAdd.Show vbModal
  If UBound(Files) = 0 Then Exit Sub
  If FileVerify(MainFile) = False Then
    ReDim Files(0)
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  Open MainFile For Binary Access Read Write As #1
  Get #1, , TmpInt
  If 32767 - UBound(Files) < TmpInt Then
    MsgBox "Archive do not support so many files", vbCritical
    Close
    Exit Sub
  End If
  TmpInt = TmpInt + UBound(Files)
  Seek #1, 1
  Put #1, , TmpInt
  Seek #1, LOF(1) + 1
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Adding files"
  For i = 1 To UBound(Files)
    FileLength = FileLength + FileLen(Files(i))
  Next i
  For i = 1 To UBound(Files)
    Open Files(i) For Binary Access Read As #2
    frmProgress.lblOperation.Caption = "Adding " + Simplificate(Files(i)) + "..."
    DoEvents
    Info.Size = LOF(2)
    Info.Modified = FileDateTime(Files(i))
    Info.Packed = 0
    If Len(MainPath) = 0 Then
      TmpStr = Simplificate(Files(i))
    Else
      If LCase(Left(Files(i), Len(MainPath))) = LCase(MainPath) Then
        TmpStr = Mid(Files(i), Len(MainPath) + 1)
      Else
        TmpStr = Simplificate(Files(i))
      End If
    End If
    ToFile 2, Chr(Len(TmpStr)) + TmpStr
    Position = Seek(1)
    Put #1, , Info
    While Loc(2) < LOF(2)
      ReDim Buff(GetSize(2))
      Get #2, , Buff()
      TmpLng = UBound(Buff)
      CompressArray Buff, Val(Files(0))
      If Len(MainPassword) <> 0 Then
        CodeArray Buff(), StrConv(MainPassword, vbFromprocode)
        ToFile 0, -UBound(Buff)
      Else
        ToFile 0, UBound(Buff)
      End If
      ToFile 0, TmpLng
      Put #1, , Buff()
      Info.Packed = Info.Packed + UBound(Buff) + 9
      frmProgress.Calculate FileLength, CurrentPos + Seek(2)
      DoEvents
    Wend
    CurrentPos = CurrentPos + LOF(2)
    ToFile 0, 0
    Info.Packed = Info.Packed + 4
    Seek #1, Position
    Put #1, , Info
    Seek #1, LOF(1) + 1
    Close #2
  Next i
  Me.Enabled = True
  Unload frmProgress
  Close
  ReDim Files(0)
  LoadFiles MainFile, False
End Sub

Private Sub mnuClose_Click()
  MainFile = ""
  lstFiles.ListItems.Clear
  Me.Caption = AppName
  lstFiles.OLEDropMode = ccOLEDropNone
End Sub

Private Sub mnuComments_Click()
  Dim TmpStr As String
  Dim Info As proInfo
  Dim i As Integer, TmpInt As Integer
  Dim Position As Long, FileLength As Long, CurrentPos As Long
  Dim Buff() As Byte

  Load frmViewer
  With frmViewer
    .txtText.Locked = False
    .Caption = "Archive comment"
    .txtText.Text = MainComment
    TmpStr = MainComment
    .Show vbModal
    If MainComment = TmpStr Then Exit Sub
  End With
  If FileVerify(MainFile) = False Then
    MainComment = TmpStr
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Refreshing comment"
  frmProgress.lblOperation.Caption = "Refreshing comment..."
  DoEvents
  Open MainFile For Binary Access Read As #2
  Open MainFile + ".tmp" For Binary Access Write As #1
  Get #2, , TmpInt
  For i = 1 To TmpInt
    FileLength = FileLength + Val(lstFiles.ListItems(i).SubItems(2))
  Next i
  Put #1, , TmpInt
  Get #2, , i
  TmpStr = Input(i, 2)
  ToFile 1, Len(MainComment)
  ToFile 2, MainComment
  For i = 1 To TmpInt
    TmpStr = Input(Asc(Input(1, 2)), 2)
    Get #2, , Info
    ToFile 2, Chr(Len(TmpStr)) + TmpStr
    Put #1, , Info
    Do
      Get #2, , Position
      CurrentPos = CurrentPos + Abs(Position) + 4
      If Position = 0 Then Exit Do
      ReDim Buff(Abs(Position))
      Put #1, , Position
      Get #2, , Position
      Put #1, , Position
      Get #2, , Buff()
      Put #1, , Buff()
      frmProgress.Calculate FileLength, CurrentPos
      DoEvents
    Loop
    ToFile 0, 0
  Next i
  Close
  Kill MainFile
  Name MainFile + ".tmp" As MainFile
  Me.Enabled = True
  Unload frmProgress
  LoadFiles MainFile, False
End Sub

Private Sub mnuContent_Click()
  WinHelp Me.hwnd, HelpFile, &HB, 0
End Sub

Private Sub mnuDelete_Click()
  Dim TmpStr As String
  Dim Info As proInfo
  Dim i As Integer, TmpInt As Integer
  Dim TmpVar As Variant
  Dim Position As Long, FileLength As Long, CurrentPos As Long
  Dim Buff() As Byte

  If Len(MainFile) = 0 Then
    MsgBox "No open archive", vbCritical
    Exit Sub
  End If
  TmpVar = GetSelected()
  If TmpVar(0) = 0 Then
    MsgBox "No files selected to delete", vbCritical
    Exit Sub
  End If
  If MsgBox("Are you sure you want to delete selected file(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
  If FileVerify(MainFile) = False Then
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Deleting files"
  frmProgress.lblOperation.Caption = "Deleting..."
  DoEvents
  Open MainFile For Binary Access Read As #2
  Open MainFile + ".tmp" For Binary Access Write As #1
  Get #2, , TmpInt
  For i = 1 To TmpInt
    If TmpVar(i) = False Then FileLength = FileLength + Val(lstFiles.ListItems(i).SubItems(2))
  Next i
  ToFile 1, TmpInt - TmpVar(0)
  Get #2, , i
  TmpStr = Input(i, 2)
  Put #1, , i
  ToFile 2, TmpStr
  For i = 1 To TmpInt
    TmpStr = Input(Asc(Input(1, 2)), 2)
    Get #2, , Info
    If TmpVar(i) = False Then
      ToFile 2, Chr(Len(TmpStr)) + TmpStr
      Put #1, , Info
      Do
        Get #2, , Position
        CurrentPos = CurrentPos + Abs(Position) + 4
        If Position = 0 Then Exit Do
        ReDim Buff(Abs(Position))
        Put #1, , Position
        Get #2, , Position
        Put #1, , Position
        Get #2, , Buff()
        Put #1, , Buff()
        frmProgress.Calculate FileLength, CurrentPos
        DoEvents
      Loop
      ToFile 0, 0
    Else
      Seek #2, Seek(2) + Info.Packed
    End If
  Next i
  Close
  Kill MainFile
  Name MainFile + ".tmp" As MainFile
  Me.Enabled = True
  Unload frmProgress
  LoadFiles MainFile, False
End Sub

Private Sub mnuExit_Click()
  End
End Sub

Private Sub mnuExtract_Click()
  Dim TmpVar As Variant
  Dim i As Integer, TmpInt As Integer
  Dim TmpStr As String, Folder As String, DecompressInfo() As String
  Dim TmpLng As Long, FileLength As Long, CurrentPos As Long
  Dim Buff() As Byte
  Dim Errors As Boolean, AskAllowed As Boolean, ErrorFound As Boolean, Encrypted As Boolean
  Dim Info As proInfo

  If Len(MainFile) = 0 Then
    MsgBox "No open archive", vbCritical
    Exit Sub
  End If
  TmpVar = GetSelected()
  If TmpVar(0) = 0 Then
    MsgBox "No files selected to extract", vbCritical
    Exit Sub
  End If
  If Len(ExtractPath) = 0 Then
    frmExtract.Show vbModal
    If Len(ExtractPath) = 0 Then Exit Sub
  Else
    AskAllowed = True
  End If
  If FileVerify(MainFile) = False Then
    ExtractPath = ""
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Extracting files"
  frmProgress.lblOperation.Caption = "Searching..."
  DoEvents
  Open MainFile For Binary Access Read As #1
  Get #1, , TmpInt
  ReDim DecompressInfo(1 To TmpInt)
  Get #1, , i
  Seek #1, Seek(1) + i
  For i = 1 To TmpInt
    If TmpVar(i) = True Then FileLength = FileLength + Val(lstFiles.ListItems(i).SubItems(1))
  Next i
  For i = 1 To TmpInt
    TmpStr = Input(Asc(Input(1, 1)), 1)
    Get #1, , Info
    If TmpVar(i) = True Then
      If Options(1) And FileVerify(ExtractPath + TmpStr) Then If FileDateTime(ExtractPath + TmpStr) > CDate(Info.Modified) Then GoTo Continue
      If Options(0) Then
        If FileVerify(ExtractPath + TmpStr) Then
          Select Case MsgBox("File '" + TmpStr + "' already exists" + vbCrLf + "Overwrite?", vbQuestion + vbYesNoCancel)
            Case vbCancel
              Close
              ExtractPath = ""
              Exit Sub
            Case vbNo
              GoTo Continue
          End Select
        End If
      End If
      If Options(2) And InStr(TmpStr, "\") Then
        Folder = Left(TmpStr, Len(TmpStr) - Len(Simplificate(TmpStr)))
        TmpStr = Simplificate(TmpStr)
        CreateDir ExtractPath, Folder
        Open ExtractPath + Folder + TmpStr For Output As #2
        Close #2
        Open ExtractPath + Folder + TmpStr For Binary Access Write As #2
      Else
        TmpStr = Simplificate(TmpStr)
        Open ExtractPath + TmpStr For Output As #2
        Close #2
        Open ExtractPath + TmpStr For Binary Access Write As #2
      End If
      frmProgress.lblOperation.Caption = "Extracting " + TmpStr + "..."
      DoEvents
      Errors = False
      Do
        Get #1, , TmpLng
        If TmpLng = 0 Then Exit Do
        Encrypted = False
        If TmpLng < 0 Then
          Encrypted = True
          TmpLng = -TmpLng
          If AskAllowed Then MainPassword = frmAskPassword.InputPassword(): Unload frmAskPassword
        End If
        ReDim Buff(TmpLng)
        Get #1, , TmpLng
        Get #1, , Buff()
        If Encrypted And Len(MainPassword) > 0 Then CodeArray Buff(), StrConv(MainPassword, vbFromprocode)
        If DecompressArray(Buff, TmpLng) <> 0 Then Errors = True
        Put #2, , Buff()
        frmProgress.Calculate FileLength, CurrentPos + Loc(2)
        DoEvents
      Loop
      CurrentPos = CurrentPos + LOF(2)
      Close #2
      If Errors Then
        ErrorFound = True
        DecompressInfo(i) = TmpStr + ": Error in decompression"
        If Encrypted Then DecompressInfo(i) = DecompressInfo(i) + ". Check password"
      Else
        DecompressInfo(i) = TmpStr + ": OK"
      End If
Continue:
    Else
      Seek #1, Seek(1) + Info.Packed
    End If
  Next i
  Close
  Me.Enabled = True
  Unload frmProgress
  If ErrorFound Then
    If MsgBox("Errors found in decompression" + vbCrLf + "Would you like to see the log file?", vbQuestion + vbYesNo) = vbYes Then
      Load frmViewer
      For i = 1 To UBound(DecompressInfo)
        If Len(DecompressInfo(i)) > 0 Then
          frmViewer.txtText = frmViewer.txtText + DecompressInfo(i)
          If i <> UBound(DecompressInfo) Then frmViewer.txtText = frmViewer.txtText + vbCrLf
        End If
      Next i
      frmViewer.Show vbModal
    End If
  End If
  ExtractPath = ""
End Sub

Private Sub mnuFile_Click()
  mnuClose.Enabled = Len(MainFile) <> 0
  mnuSearch.Enabled = mnuClose.Enabled
  mnuRename.Enabled = mnuClose.Enabled And lstFiles.ListItems.Count <> 0
End Sub

Private Sub mnuMakeExe_Click()
  frmSelfExtractor.Show vbModal
End Sub

Private Sub mnuNew_Click()
  Dim File As String

  File = Dialog(Me, 1)
  If Len(File) = 0 Then Exit Sub
  Open File For Output As #1
  Close
  Open File For Binary Access Write As #1
  ToFile 1, 0
  ToFile 1, 0
  Close
  Me.Caption = AppName + " - " + Simplificate(File)
  MainFile = File
  LoadFiles MainFile, False
End Sub

Private Sub mnuOpen_Click()
  Dim File As String

  File = Dialog(Me, 0)
  If Len(File) = 0 Then Exit Sub
  Me.Caption = AppName + " - " + Simplificate(File)
  MainFile = File
  LoadFiles MainFile, True
End Sub

Private Sub mnuRename_Click()
  Dim TmpStr As String
  Dim Info As proInfo
  Dim i As Integer, TmpInt As Integer
  Dim Position As Long, FileLength As Long, CurrentPos As Long
  Dim Buff() As Byte

  frmRename.Show vbModal
  If UBound(Files) = 0 Then Exit Sub
  If FileVerify(MainFile) = False Then
    ReDim Files(0)
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Renaming files"
  frmProgress.lblOperation.Caption = "Renaming..."
  DoEvents
  Open MainFile For Binary Access Read As #2
  Open MainFile + ".tmp" For Binary Access Write As #1
  Get #2, , TmpInt
  For i = 1 To TmpInt
    FileLength = FileLength + Val(lstFiles.ListItems(i).SubItems(2))
  Next i
  Put #1, , TmpInt
  Get #2, , i
  TmpStr = Input(i, 2)
  Put #1, , i
  ToFile 2, TmpStr
  For i = 1 To TmpInt
    TmpStr = Input(Asc(Input(1, 2)), 2)
    Get #2, , Info
    Files(i) = lstFiles.ListItems(i).SubItems(5) + Files(i)
    ToFile 2, Chr(Len(Files(i))) + Files(i)
    Put #1, , Info
    Do
      Get #2, , Position
      CurrentPos = CurrentPos + Abs(Position) + 4
      If Position = 0 Then Exit Do
      ReDim Buff(Abs(Position))
      Put #1, , Position
      Get #2, , Position
      Put #1, , Position
      Get #2, , Buff()
      Put #1, , Buff()
      frmProgress.Calculate FileLength, CurrentPos
      DoEvents
    Loop
    ToFile 0, 0
  Next i
  Close
  ReDim Files(0)
  Kill MainFile
  Name MainFile + ".tmp" As MainFile
  Me.Enabled = True
  Unload frmProgress
  LoadFiles MainFile, False
End Sub

Private Sub mnuSearch_Click()
  frmSearchFile.Show vbModal
  lstFiles_MouseUp 1, 0, 0, 0
End Sub

Private Sub mnuTest_Click()
  Dim TmpStr As String, Errors() As String
  Dim ErrRet As Integer, i As Integer, TmpInt As Integer
  Dim FileLength As Long, CurrentPos As Long, TmpLng As Long
  Dim Info As proInfo
  Dim ErrFound As Boolean, Encrypted As Boolean, Asked As Boolean
  Dim Buff() As Byte

  If Len(MainFile) = 0 Then
    MsgBox "No open archive", vbCritical
    Exit Sub
  End If
  If lstFiles.ListItems.Count = 0 Then
    MsgBox "No files to test", vbCritical
    Exit Sub
  End If
  If FileVerify(MainFile) = False Then
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  ExtractPath = Environ("temp")
  If Right(ExtractPath, 1) <> "\" Then ExtractPath = ExtractPath + "\"
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Testing files"
  DoEvents
  MainPassword = ""
  Open MainFile For Binary Access Read As #1
  Get #1, , TmpInt
  ReDim Errors(1 To TmpInt)
  For i = 1 To TmpInt
    FileLength = FileLength + Val(lstFiles.ListItems(i).SubItems(1))
  Next i
  Get #1, , i
  Seek #1, Seek(1) + i
  For i = 1 To TmpInt
    TmpStr = Simplificate(Input(Asc(Input(1, 1)), 1))
    Get #1, , Info
    frmProgress.lblOperation.Caption = "Testing " + TmpStr + "..."
    DoEvents
    Open ExtractPath + TmpStr For Output As #2
    Close #2
    Open ExtractPath + TmpStr For Binary Access Write As #2
    ErrRet = 0
    Do
      Get #1, , TmpLng
      If TmpLng = 0 Then Exit Do
      Encrypted = False
      If TmpLng < 0 Then
        If Len(MainPassword) = 0 And Asked = False Then MainPassword = frmAskPassword.InputPassword(): Unload frmAskPassword
        TmpLng = -TmpLng
        Encrypted = True
        Asked = True
      End If
      ReDim Buff(TmpLng)
      Get #1, , TmpLng
      Get #1, , Buff()
      If Encrypted And Len(MainPassword) > 0 Then CodeArray Buff(), StrConv(MainPassword, vbFromprocode)
      ErrRet = DecompressArray(Buff, TmpLng)
      Put #2, , Buff()
      frmProgress.Calculate FileLength, CurrentPos + Seek(2)
      DoEvents
    Loop
    CurrentPos = CurrentPos + LOF(2)
    Close #2
    Kill ExtractPath + TmpStr
    If ErrRet <> 0 Then ErrFound = True
    Select Case ErrRet
      Case 0
        Errors(i) = TmpStr + ": OK"
      Case 1
        Errors(i) = TmpStr + ": Unexpected end of file found"
      Case 2
        Errors(i) = TmpStr + ": Cannot find dictionary"
      Case -2
        Errors(i) = TmpStr + ": Invalid compressed data"
      Case -6
        Errors(i) = TmpStr + ": Wrong version"
      Case -4
        Errors(i) = TmpStr + ": Memory error"
      Case Else
        Errors(i) = TmpStr + ": Unknown error"
    End Select
    If Encrypted And ErrRet <> 0 Then Errors(i) = Errors(i) + ". Check password"
  Next i
  Close
  Me.Enabled = True
  Unload frmProgress
  Load frmViewer
  If ErrFound = True Then
     frmViewer.txtText = "Some errors were found" + vbCrLf + vbCrLf
  Else
     frmViewer.txtText = "No errors were found" + vbCrLf + vbCrLf
  End If
  For i = 1 To UBound(Errors)
    frmViewer.txtText = frmViewer.txtText + Errors(i)
    If i <> UBound(Errors) Then frmViewer.txtText = frmViewer.txtText + vbCrLf
  Next i
  frmViewer.Show vbModal
  ExtractPath = ""
End Sub

Private Sub mnuViewer_Click()
  Dim TmpStr As String
  Dim i As Integer

  If Len(MainFile) = 0 Then Exit Sub
  ExtractPath = Environ("temp")
  If Right(ExtractPath, 1) <> "\" Then ExtractPath = ExtractPath + "\"
  TmpStr = ExtractPath
  mnuExtract_Click
  ExtractPath = TmpStr
  For i = 1 To lstFiles.ListItems.Count
    TmpStr = ExtractPath + lstFiles.ListItems(i).Text
    If lstFiles.ListItems(i).Selected And FileVerify(TmpStr) Then
      Open TmpStr For Binary Access Read As #1
      Load frmViewer
      frmViewer.Caption = frmViewer.Caption + " - " + lstFiles.ListItems(i).Text
      frmViewer.txtText.Text = Input(LOF(1), 1)
      Close
      frmViewer.Show vbModal
      If FileVerify(TmpStr) = True Then Kill TmpStr
    End If
  Next i
  ExtractPath = ""
End Sub

Private Sub tlbStandard_ButtonClick(ByVal Button As Button)
  Select Case Button.Index
    Case 1
      mnuNew_Click
    Case 2
      mnuOpen_Click
    Case 4
      mnuAdd_Click
    Case 5
      mnuExtract_Click
    Case 6
      mnuDelete_Click
    Case 8
      mnuTest_Click
  End Select
End Sub

Sub LoadFiles(FileName As String, ShowComment As Boolean)
  Dim TmpInt As Integer, i As Integer
  Dim Info As proInfo
  Dim TmpStr As String
  Dim TotalSize As Long

  lstFiles.ListItems.Clear
  Open FileName For Binary Access Read As #1
  Get #1, , TmpInt
  Get #1, , i
  MainComment = Input(i, 1)
  For i = 1 To TmpInt
    TmpStr = Input(Asc(Input(1, 1)), 1)
    If InStr(TmpStr, "\") <> 0 Then
      lstFiles.ListItems.Add , , Simplificate(TmpStr)
      lstFiles.ListItems(i).SubItems(5) = Left(TmpStr, Len(TmpStr) - Len(Simplificate(TmpStr)))
    Else
      lstFiles.ListItems.Add , , TmpStr
    End If
    'lstFiles.ListItems.Add , , Input(Asc(Input(1, 1)), 1)
    Get #1, , Info
    TotalSize = TotalSize + Info.Size
    lstFiles.ListItems(i).SubItems(1) = Format(Info.Size) + " bytes"
    lstFiles.ListItems(i).SubItems(2) = Format(Info.Packed) + " bytes"
    If Info.Size = 0 Or Info.Size <= Info.Packed Then
      lstFiles.ListItems(i).SubItems(3) = "0 %"
    Else
      lstFiles.ListItems(i).SubItems(3) = Format(Int(100 - Info.Packed / Info.Size * 100)) + " %"
    End If
    lstFiles.ListItems(i).SubItems(4) = CDate(Info.Modified)
    Seek #1, Seek(1) + Info.Packed
  Next i
  Close
  stbStatus.Panels(1).Text = Format(Int(TotalSize / 1024)) + " KB"
  stbStatus.Panels(2).Text = "0 KB"
  stbStatus.Panels(3).Text = Format(lstFiles.ListItems.Count) + " file(s)"
  stbStatus.Panels(4).Text = "0 file(s) selected"
  lstFiles.OLEDropMode = ccOLEDropManual
  If Len(MainComment) <> 0 And ShowComment Then
    Load frmViewer
    frmViewer.Caption = "Archive comment"
    frmViewer.txtText.Text = MainComment
    frmViewer.Show vbModal
  End If
End Sub

Function GetSelected() As Variant
  Dim i As Integer, TmpInt As Integer
  Dim TmpVar() As Variant

  ReDim TmpVar(lstFiles.ListItems.Count)
  For i = 1 To UBound(TmpVar)
    TmpVar(i) = lstFiles.ListItems(i).Selected
    If TmpVar(i) = True Then TmpInt = TmpInt + 1
  Next i
  TmpVar(0) = TmpInt
  GetSelected = TmpVar
End Function

Sub CreateDir(ByVal CurrentPath As String, NewPath As String)
  Dim TmpStr() As String
  Dim i As Integer

  If ExistDir(CurrentPath + NewPath) Then Exit Sub
  TmpStr = Split(Left(NewPath, Len(NewPath) - 1), "\")
  For i = 0 To UBound(TmpStr)
    If Not ExistDir(CurrentPath + TmpStr(i)) Then MkDir CurrentPath + TmpStr(i)
    CurrentPath = CurrentPath + TmpStr(i) + "\"
  Next i
End Sub

Function ExistDir(DirPath As String) As Boolean
  On Error GoTo Verificate
  ChDir DirPath
  ExistDir = True
  Exit Function
Verificate:
  ExistDir = False
End Function

Sub UpdateFile(FileNum As Integer, NewFileName As String, NewPath As String)
  Dim Info As proInfo
  Dim TmpInt As Integer, i As Integer
  Dim TmpStr As String
  Dim Buff() As Byte
  Dim TmpLng As Long, Position As Long, FileLength As Long, CurrentPos As Long

  If FileVerify(MainFile) = False Then
    MsgBox "Archive does not exists", vbCritical
    Exit Sub
  End If
  frmProgress.Show
  Me.Enabled = False
  frmProgress.Caption = "Updating"
  frmProgress.lblOperation.Caption = "Updating " + lstFiles.ListItems(FileNum) + "..."
  DoEvents
  Open MainFile For Binary Access Read As #2
  Open MainFile + ".tmp" For Binary Access Write As #1
  Get #2, , TmpInt
  For i = 1 To TmpInt
    If i <> FileNum Then FileLength = FileLength + Val(lstFiles.ListItems(i).SubItems(2))
  Next i
  FileLength = FileLength + FileLen(NewFileName)
  ToFile 1, TmpInt
  Get #2, , i
  TmpStr = Input(i, 2)
  Put #1, , i
  ToFile 2, TmpStr
  For i = 1 To TmpInt
    TmpStr = Input(Asc(Input(1, 2)), 2)
    Get #2, , Info
    If i <> FileNum Then
      ToFile 2, Chr(Len(TmpStr)) + TmpStr
      Put #1, , Info
      Do
        Get #2, , Position
        CurrentPos = CurrentPos + Abs(Position) + 4
        If Position = 0 Then Exit Do
        ReDim Buff(Abs(Position))
        Put #1, , Position
        Get #2, , Position
        Put #1, , Position
        Get #2, , Buff()
        Put #1, , Buff()
        frmProgress.Calculate FileLength, CurrentPos
        DoEvents
      Loop
      ToFile 0, 0
    Else
      Seek #2, Seek(2) + Info.Packed
      Open NewFileName For Binary Access Read As #3
      TmpStr = NewPath + Simplificate(NewFileName)
      ToFile 2, Chr(Len(TmpStr)) + TmpStr
      Position = Seek(1)
      Info.Size = LOF(3)
      Info.Modified = FileDateTime(NewFileName)
      Info.Packed = 0
      Put #1, , Info
      While Loc(3) < LOF(3)
        ReDim Buff(GetSize(3))
        Get #3, , Buff()
        TmpLng = UBound(Buff)
        CompressArray Buff, 9
        If Len(MainPassword) <> 0 Then
          CodeArray Buff(), StrConv(MainPassword, vbFromprocode)
          ToFile 0, -UBound(Buff)
        Else
          ToFile 0, UBound(Buff)
        End If
        ToFile 0, TmpLng
        Put #1, , Buff()
        Info.Packed = Info.Packed + UBound(Buff) + 9
        frmProgress.Calculate FileLength, CurrentPos + Seek(3)
        DoEvents
      Wend
      CurrentPos = CurrentPos + LOF(3)
      ToFile 0, 0
      Info.Packed = Info.Packed + 4
      Seek #1, Position
      Put #1, , Info
      Seek #1, LOF(1) + 1
      Close #3
    End If
  Next i
  Close
  Kill MainFile
  Name MainFile + ".tmp" As MainFile
  Me.Enabled = True
  Unload frmProgress
  LoadFiles MainFile, False
End Sub

