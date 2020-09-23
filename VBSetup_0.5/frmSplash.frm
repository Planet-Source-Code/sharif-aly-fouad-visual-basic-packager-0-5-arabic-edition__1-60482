VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4110
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   2160
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   3720
   End
   Begin VB.Timer timFireScroll 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   120
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5640
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picScroolContainer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         Height          =   2880
         Left            =   240
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   2
         Top             =   240
         Width           =   5175
         Begin VB.PictureBox picCredits 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   4095
            Left            =   -120
            ScaleHeight     =   273
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   345
            TabIndex        =   3
            Top             =   -1440
            Width           =   5175
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   $"frmSplash.frx":000C
               Height          =   390
               Left            =   2400
               TabIndex        =   10
               Top             =   2760
               Width           =   1305
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "æ íãßäß ÇíÖÇ ÇáÇÔÊÑÇß ãÚäÇ ÈÇí ÇÖÇÝå ááãÔÑæÚ"
               Height          =   195
               Left            =   1080
               TabIndex        =   9
               Top             =   3240
               Width           =   3225
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "ÈÇáÇÔÊÑÇß ãÚ : ÔÑíÝ Úáí"
               Height          =   195
               Left            =   1800
               TabIndex        =   6
               Top             =   2280
               Width           =   1755
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "ÈÑãÌå : ÇÈæ ÚíÏ ÇÈÑÇåíã"
               Height          =   195
               Left            =   1800
               TabIndex        =   5
               Top             =   1320
               Width           =   1515
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ãæÞÚ ÝíÌæÇá ÈíÓß ááÚÑÈ "
               Height          =   195
               Left            =   1800
               TabIndex        =   4
               Top             =   360
               Width           =   1785
            End
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&O K"
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   3240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide
'timFireScroll.Enabled = False
End Sub

Private Sub Form_Load()

timFireScroll.Enabled = True
Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub Label2_Click()

End Sub

Private Sub picCredits_Click()

End Sub

Private Sub Timer1_Timer()
Dim x
x = Text1.Text + 1
Text1.Text = x

If x = 140 Then x = 0: Text1.Text = "00": Me.Hide
'picCredits.Top = picCredits.Top + 25:


End Sub

Private Sub Timer2_Timer()
'DoEvents
Text2.Text = picCredits.Top
picCredits.Top = picCredits.Top + 390
If Text2.Text <= 390 Then Timer2.Enabled = False: timFireScroll.Enabled = True

End Sub

Private Sub timFireScroll_Timer()
'Dim c
Text2.Text = picCredits.Top

DoEvents
picCredits.Top = picCredits.Top - 1
If Text2.Text <= -212 Then timFireScroll.Enabled = False: Timer2.Enabled = True

End Sub

