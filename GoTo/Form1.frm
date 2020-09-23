VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GoTo! - "
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox LoadHome 
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   230
      Left            =   5040
      TabIndex        =   13
      Top             =   400
      Width           =   975
   End
   Begin VB.TextBox txtHome 
      Height          =   285
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox txtOperation 
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   8805
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   75
      Left            =   720
      TabIndex        =   4
      Top             =   1365
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Max             =   4
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GoTo!"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox url 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "http://www.google.com"
      Top             =   1080
      Width           =   6615
   End
   Begin SHDocVwCtl.WebBrowser browse 
      Height          =   7320
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   9615
      ExtentX         =   16960
      ExtentY         =   12912
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "Search"
      Height          =   255
      Left            =   3640
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Image imgSearch 
      Height          =   630
      Left            =   3600
      Picture         =   "Form1.frx":0082
      ToolTipText     =   "Search"
      Top             =   190
      Width           =   630
   End
   Begin VB.Label Label5 
      Caption         =   "Home"
      Height          =   255
      Left            =   4420
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgHome 
      Height          =   630
      Left            =   4320
      Picture         =   "Form1.frx":15C4
      ToolTipText     =   "Home"
      Top             =   195
      Width           =   630
   End
   Begin VB.Label Label4 
      Caption         =   "Forward"
      Height          =   255
      Left            =   2900
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Image imgNext 
      Height          =   630
      Left            =   2880
      Picture         =   "Form1.frx":2B06
      ToolTipText     =   "Forward"
      Top             =   190
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "Back"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgPrev 
      Height          =   630
      Left            =   2160
      Picture         =   "Form1.frx":4048
      ToolTipText     =   "Back"
      Top             =   195
      Width           =   630
   End
   Begin VB.Label Label3 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   1475
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   880
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgRef 
      Height          =   630
      Left            =   1440
      Picture         =   "Form1.frx":558A
      ToolTipText     =   "Refresh"
      Top             =   190
      Width           =   630
   End
   Begin VB.Image imgStop 
      Height          =   630
      Left            =   720
      Picture         =   "Form1.frx":6ACC
      ToolTipText     =   "Stop"
      Top             =   190
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1125
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Integer
Dim o As Integer
Dim ErrHandle

Private Sub bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
    .Panels.Item(1).Text = ""
End With

End Sub

Private Sub browse_BeforeNavigate2(ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If browse.ReadyState = READYSTATE_LOADING Then
    Form1.Caption = "GoTo! - " & browse.LocationURL
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    Form1.Caption = "GoTo! - " & browse.LocationName
End If
End Sub

Private Sub browse_DocumentComplete(ByVal pDisp As Object, url As Variant)
'If browse.ReadyState = READYSTATE_LOADING Then
'    Form1.Caption = "GoTo! - " & browse.LocationURL
'End If
'If browse.ReadyState = READYSTATE_COMPLETE Then
'    Form1.Caption = "GoTo! - " & browse.LocationName
'End If
'If browse.ReadyState = READYSTATE_LOADING Then
'    bar.Value = browse.ReadyState
'End If
'If browse.ReadyState = READYSTATE_COMPLETE Then
'    statusbar.Panels.Item(1).Text = "Complete"
'End If
'If browse.ReadyState = READYSTATE_LOADED Then
'    bar.Value = 0
'End If
End Sub

Private Sub browse_DownloadBegin()
'statusbar.Panels.Item(1).Text = "Downloading File..."
End Sub

Private Sub browse_DownloadComplete()
'statusbar.Panels.Item(1).Text = "Download Complete"
End Sub

Private Sub browse_NavigateComplete2(ByVal pDisp As Object, url As Variant)
If browse.ReadyState = READYSTATE_LOADING Then
    Form1.Caption = "GoTo! - " & browse.LocationURL
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    Form1.Caption = "GoTo! - " & browse.LocationName
End If
If browse.ReadyState = READYSTATE_LOADING Then
    bar.Value = browse.ReadyState
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    statusbar.Panels.Item(1).Text = "Complete"
End If
If browse.ReadyState = READYSTATE_LOADED Then
    bar.Value = 0
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    lblStop.Enabled = False
    imgStop.Enabled = False
End If
If browse.ReadyState = READYSTATE_LOADING Then
    lblStop.Enabled = True
    imgStop.Enabled = True
End If

End Sub

Private Sub browse_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
bar.Value = browse.ReadyState
If browse.ReadyState = READYSTATE_LOADING Then
    statusbar.Panels.Item(1).Text = "Transferring data..."
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    bar.Value = 0
    statusbar.Panels.Item(1).Text = "Complete"
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    Form1.Caption = "GoTo! - " & browse.LocationName
    url.Text = browse.LocationURL
End If
If browse.ReadyState = READYSTATE_LOADING Then
    Form1.Caption = "GoTo! - " & browse.LocationURL
    url.Text = browse.LocationURL
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    lblStop.Enabled = False
    imgStop.Enabled = False
End If
If browse.ReadyState = READYSTATE_LOADING Then
    lblStop.Enabled = True
    imgStop.Enabled = True
End If
End Sub

Private Sub browse_TitleChange(ByVal Text As String)
If browse.ReadyState = READYSTATE_LOADING Then
    url.Text = browse.LocationURL
    Form1.Caption = browse.LocationName
End If
If browse.ReadyState = READYSTATE_COMPLETE Then
    Form1.Caption = browse.LocationName
    url.Text = browse.LocationURL
End If

End Sub

Private Sub Command1_Click()
browse.Navigate (url.Text)
o = 0
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
    .Panels.Item(1).Text = "Click here to GoTo your destination"
End With

End Sub



Private Sub Command2_Click()
frmHome.Visible = True
End Sub

Private Sub Form_Load()
LoadHome.LoadFile ("c:\gotoprefs.pre")
txtHome.Text = LoadHome.Text
statusbar.Panels.Item(1).Width = 9600
browse.Navigate "http://www.google.com"
url.Text = "http://www.google.com"
statusbar.Panels.Item(1).Text = "Waiting..."
imgHome.ToolTipText = "Home " & txtHome.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub


Private Sub imgHome_Click()
browse.Navigate (txtHome.Text)
End Sub

Private Sub imgHome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHome.Picture = LoadPicture("c:\homedown.bmp")
End Sub

Private Sub imgHome_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHome.Picture = LoadPicture("c:\home.bmp")

End Sub

Private Sub imgNext_Click()
browse.GoForward
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = LoadPicture("c:\nextdown.bmp")
End Sub

Private Sub imgNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNext.Picture = LoadPicture("c:\next.bmp")
End Sub

Private Sub imgPrev_Click()
browse.GoBack

End Sub

Private Sub imgPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPrev.Picture = LoadPicture("c:\prevdown.bmp")
End Sub

Private Sub imgPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub

Private Sub imgPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPrev.Picture = LoadPicture("c:\prev.bmp")
End Sub

Private Sub imgRef_Click()
browse.Refresh
End Sub

Private Sub imgRef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRef.Picture = LoadPicture("c:\refreshdown.bmp")
End Sub

Private Sub imfRef_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With

End Sub

Private Sub imgRef_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub

Private Sub imgRef_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRef.Picture = LoadPicture("c:\refresh.bmp")
End Sub

Private Sub imgRe_Click()

End Sub

Private Sub imgSearch_Click()
browse.GoSearch
End Sub

Private Sub imgSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSearch.Picture = LoadPicture("c:\searchdown.bmp")
End Sub

Private Sub imgSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub

Private Sub imgSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSearch.Picture = LoadPicture("c:\search.bmp")
End Sub

Private Sub imgStop_Click()
If browse.ReadyState = READYSTATE_LOADING Then
    lblStop.Enabled = True
    browse.Stop
    txtOperation.Text = 1
    o = txtOperation.Text
    statusbar.Panels.Item(1).Text = "Operation Halted"
Else
    MsgBox "Cannot halt activity, no data being transferred...", vbCritical, "Cannot halt activity..."
    lblStop.Enabled = False
    imgStop.Enabled = False
End If
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStop.Picture = LoadPicture("c:\stopdown.bmp")
End Sub

Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStop.Picture = LoadPicture("c:\stop.bmp")

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If o = 1 Then
            statusbar.Panels.Item(1).Text = "Operation Halted"
        End If
End With
End Sub

Private Sub lblStatus_Click()
With statusbar
        If browse.ReadyState = READYSTATE_COMPLETE Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
        If browse.ReadyState = READYSTATE_LOADING Then
            statusbar.Panels.Item(1).Text = "Complete"
        End If
End With
End Sub

Private Sub url_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With statusbar
    .Panels.Item(1).Text = "Enter a URL here"
End With
End Sub
