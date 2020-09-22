VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   4680
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   1650
      ScaleWidth      =   4800
      TabIndex        =   17
      Top             =   6000
      Width           =   4800
      Visible         =   0   'False
      Begin VB.CommandButton Command6 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   360
         Picture         =   "Form1.frx":1A58C
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "WinLock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   75
         Width           =   2775
      End
      Begin VB.Label Labelinfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   15
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   0
      Picture         =   "Form1.frx":1A703
      ScaleHeight     =   2700
      ScaleWidth      =   4800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3360
      Width           =   4800
      Visible         =   0   'False
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "Form1.frx":44A45
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Change password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   75
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter the old password and the new password in the boxes below. Please note that the changes will take place immedeately."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current password:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New password:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm password:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   0
      Picture         =   "Form1.frx":44F38
      ScaleHeight     =   2100
      ScaleWidth      =   4800
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4800
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   360
         Picture         =   "Form1.frx":65C7A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   22
         Top             =   480
         Width           =   480
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Shutdown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   75
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This computer is locked by WinLock. You will need an administrator password to access this computer."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4920
      Picture         =   "Form1.frx":668BC
      Top             =   6360
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   5520
      Picture         =   "Form1.frx":66A33
      Top             =   6360
      Width           =   495
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pwd As String
Dim tries As Integer
Dim changed As Boolean
Dim dn As Boolean
Private Sub Command1_Click()
RestartWindows False, False

End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text1.Text > "" Then
Command1.Enabled = False
Command3.Enabled = False
Command2.Enabled = False
pwd = GetSetting("secur", "login", "pwd", "guest")
If Text1.Text = pwd Then
DisableKeys (False)
End
Else
Beep
Labelinfo.Caption = "You have entered a invalid password! Do not try to login on this computer if you aren't authorized!"
Picture3.Visible = True
Image3.Picture = Image5.Picture
Text1.Text = ""
Text1.SetFocus
Command6.SetFocus
End If
End If
Text1.SetFocus
Command6.SetFocus
End Sub

Private Sub Command3_Click()
Picture2.Visible = True
Text2.SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next
Command4.Enabled = False
Command5.Enabled = False
Command6.SetFocus
If GetSetting("secur", "login", "pwd", "guest") = Text2.Text Then
GoTo ok
Else
Labelinfo.Caption = "You have entered a invalid password! Do not try to login on this computer if you aren't authorized!"
Image3.Picture = Image5.Picture
Picture3.Visible = True
Beep
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command6.SetFocus
Exit Sub
End If
ok:
If Text3.Text = Text4.Text Then
If Text3.Text = "" Then GoTo nepp
Picture3.Visible = True
Beep
Image3.Picture = Image4.Picture
Labelinfo.Caption = "You have successfully changed the password!"
changed = True
SaveSetting "secur", "login", "pwd", Text3.Text
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command6.SetFocus
Exit Sub
Else
Picture3.Visible = True
Beep
Image3.Picture = Image5.Picture
Labelinfo.Caption = "The password doesn't match!"
Text3.Text = ""
Text4.Text = ""
Text3.SetFocus
Command6.SetFocus
Exit Sub
End If
nepp:
Picture3.Visible = True
Beep
Image3.Picture = Image5.Picture
Labelinfo.Caption = "You must specify a password!"
Text3.Text = ""
Text4.Text = ""
Text3.SetFocus
Command6.SetFocus
End Sub

Private Sub Command5_Click()
Picture2.Visible = False
Text1.SetFocus
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command5_LostFocus()
On Error Resume Next
Text2.SetFocus
End Sub

Private Sub Command6_Click()
On Error Resume Next
If changed = True Then
Picture2.Visible = False
changed = False
Text1.SetFocus
End If
If Labelinfo.Caption = "You must specify a password!" Then
Text3.SetFocus
End If
If Labelinfo.Caption = "The password doesn't match!" Then
Text3.SetFocus
End If
Picture3.Visible = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command5.Enabled = True
Command4.Enabled = True

End Sub

Private Sub Form_Activate()
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, conSwpNoActivate Or conSwpShowWindow
Text1.SetFocus
End Sub
Private Sub Form_Load()
App.TaskVisible = False
DisableKeys (True)
Me.Width = Screen.Width
Me.Height = Screen.Height
Picture1.Left = Me.ScaleWidth / 2 - Picture1.ScaleWidth / 2
Picture1.Top = Me.ScaleHeight / 2.5 - Picture1.ScaleHeight
Picture2.Left = Me.ScaleWidth / 2 - Picture1.ScaleWidth / 2
Picture2.Top = Me.ScaleHeight / 2.5 - Picture1.ScaleHeight
Picture3.Left = Me.ScaleWidth / 2 - Picture1.ScaleWidth / 2
Picture3.Top = Me.ScaleHeight / 2.5 - Picture1.ScaleHeight
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, conSwpNoActivate Or conSwpShowWindow

End Sub


Private Sub Image1_Click()
End
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Text1.Text > "" And KeyCode = "13" Then
Command2_Click
End If
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = "13" Then
Command4_Click
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = "13" Then
Command4_Click
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = "13" Then
Command4_Click
End If
End Sub

Private Sub Timer1_Timer()
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, conSwpNoActivate Or conSwpShowWindow

End Sub
