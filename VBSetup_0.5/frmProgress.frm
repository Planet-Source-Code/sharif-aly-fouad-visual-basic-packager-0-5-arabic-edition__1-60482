VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   1170
   ClientTop       =   1545
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   975
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblOperation 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Calculate(Val1 As Long, Val2 As Long)
  Dim TmpInt As Integer

  TmpInt = Int(Val2 / Val1 * 100)
  If Val1 < Val2 Or prgProgress.Value = TmpInt Then Exit Sub
  prgProgress.Value = TmpInt
End Sub
