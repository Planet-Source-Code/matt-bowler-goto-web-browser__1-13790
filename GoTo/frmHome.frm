VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GoTo! - Set Homepage"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Home 
      Height          =   310
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmHome.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change!"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open "c:\gotoprefs.pre" For Output As #1
Print #1, Home.Text
Close #1
Form1.txtHome.Text = Home.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Home.LoadFile ("c:\gotoprefs.pre")
End Sub
