Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Compress Lib "ppacklib" Alias "compress2" (Dest As Any, DestLen As Any, Src As Any, ByVal SrcLen As Long, ByVal Level As Long) As Long
Private Declare Function Decompress Lib "ppacklib" Alias "uncompress" (Dest As Any, DestLen As Any, Src As Any, ByVal SrcLen As Long) As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public MainFile As String, MainComment As String, MainPath As String, MainPassword As String, Files() As String, ExtractPath As String, HelpFile As String
Public Options(2) As Boolean
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * 260
  cAlternate As String * 14
End Type
Type proInfo
  Packed As Long
  Size As Long
  Modified As Double
End Type

Sub CreateDll()
  Dim FileNumber As Integer
  Dim DllBuffer() As Byte
  Dim TmpDir As String

  TmpDir = Environ("windir")
  If Right(TmpDir, 1) <> "\" Then TmpDir = TmpDir + "\"
  TmpDir = TmpDir + "System\ppacklib.dll"
  If FileVerify(TmpDir) Then Exit Sub
  DllBuffer = LoadResData(2, "CUSTOM")
  Open TmpDir For Binary Access Write As #1
  Put #1, , DllBuffer
  Close
End Sub

Sub CodeArray(Buff() As Byte, Password() As Byte)
  Dim i As Long
  Dim CountVar As Integer

  For i = 0 To UBound(Buff)
    Buff(i) = Buff(i) Xor Password(CountVar)
    CountVar = CountVar + 1
    If CountVar > UBound(Password) Then CountVar = 0
  Next i
End Sub

Function CompressArray(TheData() As Byte, CompressionLevel As Integer) As Long
  Dim Result As Long
  Dim BufferSize As Long
  Dim TempBuffer() As Byte

  BufferSize = UBound(TheData) + 1
  BufferSize = BufferSize + (BufferSize * 0.01) + 12
  ReDim TempBuffer(BufferSize)
  Result = Compress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)
  ReDim Preserve TheData(BufferSize - 1)
  CopyMemory TheData(0), TempBuffer(0), BufferSize
  Erase TempBuffer
  CompressArray = Result
End Function

Function DecompressArray(TheData() As Byte, OrigSize As Long) As Long
  Dim Result As Long
  Dim BufferSize As Long
  Dim TempBuffer() As Byte

  BufferSize = OrigSize
  BufferSize = BufferSize + (BufferSize * 0.01) + 12
  ReDim TempBuffer(BufferSize)
  Result = Decompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)
  ReDim Preserve TheData(BufferSize - 1)
  CopyMemory TheData(0), TempBuffer(0), BufferSize
  DecompressArray = Result
End Function

Function Dialog(Formulario As Form, Action As Integer) As String
  With Formulario.dlgDialog
    .InitDir = App.Path
    .FileName = ""
    .MaxFileSize = 10000
    Select Case Action
      Case 0
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        .Filter = "pro compressed archive (*.pro)|*.pro|All files (*.*)|*.*"
        .DialogTitle = "Open archive"
        .ShowOpen
      Case 1
        .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
        .Filter = "pro compressed archive (*.pro)|*.pro"
        .DialogTitle = "New archive"
        .ShowSave
      Case 2
        .Flags = cdlOFNFileMustExist + cdlOFNAllowMultiselect + cdlOFNHideReadOnly + cdlOFNExplorer
        .Filter = "All files (*.*)|*.*"
        .DialogTitle = "Add files"
        .ShowOpen
    End Select
    Dialog = .FileName
  End With
End Function

Sub ToFile(Tipo As Integer, Valor As Variant)
  Dim TmpLng As Long
  Dim TmpInt As Integer
  Dim TmpStr As String

  Select Case Tipo
    Case 0
      TmpLng = Valor
      Put #1, , TmpLng
    Case 1
      TmpInt = Valor
      Put #1, , TmpInt
    Case 2
      TmpStr = Valor
      Put #1, , TmpStr
  End Select
End Sub

Sub FindFiles(Path As String, Mask As String, ObjetoAdd As Object, Objeto As Object)
  Dim FindInfo As WIN32_FIND_DATA
  Dim TmpFile As String
  Dim Handler As Long, Result As Long

  If Right(Path, 1) <> "\" Then Path = Path + "\"
  Handler = FindFirstFile(Path + Mask, FindInfo)
  If Handler = -1 Then Exit Sub
  Objeto.Caption = "Scanning " + Path + "..."
  DoEvents
  Do
    TmpFile = Left(FindInfo.cFileName, InStr(FindInfo.cFileName, vbNullChar) - 1)
    If Not GetFileAttributes(Path + TmpFile) And vbDirectory Then ObjetoAdd.AddItem Path + TmpFile
    Result = FindNextFile(Handler, FindInfo)
  Loop Until Result = 0
  FindClose Handler
End Sub

Sub FindDirectories(Path As String, Mask As String, ObjetoAdd As Object, Objeto As Object)
  Dim FindInfo As WIN32_FIND_DATA
  Dim TmpFile As String
  Dim Handler As Long, Result As Long

  If Right(Path, 1) <> "\" Then Path = Path + "\"
  FindFiles Path, Mask, ObjetoAdd, Objeto
  Handler = FindFirstFile(Path + "*.*", FindInfo)
  If Handler = -1 Then Exit Sub
  Objeto.Caption = "Scanning " + Path + "..."
  DoEvents
  Do
    TmpFile = Left(FindInfo.cFileName, InStr(FindInfo.cFileName, vbNullChar) - 1)
    If TmpFile <> "." And TmpFile <> ".." Then If GetFileAttributes(Path + TmpFile) And vbDirectory Then FindDirectories Path + TmpFile, Mask, ObjetoAdd, Objeto
    Result = FindNextFile(Handler, FindInfo)
  Loop Until Result = 0
  FindClose Handler
End Sub

Sub Search(Path As String, Mask As String, ObjetoAdd As Object, Objeto As Object)
  Dim FindInfo As WIN32_FIND_DATA
  Dim TmpFile As String
  Dim Handler As Long, Result As Long

  If Right(Path, 1) <> "\" Then Path = Path + "\"
  FindFiles Path, Mask, ObjetoAdd, Objeto
  Handler = FindFirstFile(Path + "*.*", FindInfo)
  If Handler <> -1 Then
    Do
      TmpFile = Left(FindInfo.cFileName, InStr(FindInfo.cFileName, vbNullChar) - 1)
      If LCase(TmpFile) = "archivos de programa" Then Debug.Assert True
      If TmpFile <> "." And TmpFile <> ".." Then If GetAttr(Path + TmpFile) And vbDirectory Then FindDirectories Path + TmpFile, Mask, ObjetoAdd, Objeto
      Result = FindNextFile(Handler, FindInfo)
    Loop Until Result = 0
  End If
  FindClose Handler
End Sub

Function Simplificate(ByVal Cadena As String) As String
  While InStr(Cadena, "\")
    Cadena = Mid(Cadena, InStr(Cadena, "\") + 1)
  Wend
  Simplificate = Cadena
End Function

Function GetSize(FileNum As Integer) As Long
  If Loc(FileNum) + 262144 < LOF(FileNum) Then
    GetSize = 262144
  Else
    GetSize = LOF(FileNum) - Loc(FileNum) - 1
  End If
End Function

Function FileVerify(Cadena As String) As Boolean
  On Error GoTo Verificate
  FileLen Cadena
  FileVerify = True
  Exit Function
Verificate:
  FileVerify = False
End Function

