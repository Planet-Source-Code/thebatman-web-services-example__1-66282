VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Web Services"
   ClientHeight    =   7995
   ClientLeft      =   2640
   ClientTop       =   3105
   ClientWidth     =   6585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6585
   Begin VB.PictureBox picForm 
      Height          =   4800
      Left            =   2040
      ScaleHeight     =   4740
      ScaleWidth      =   3150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   705
      Width           =   3210
      Begin VB.Frame fraMain 
         Caption         =   "fraMain"
         Height          =   4095
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2775
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton cmdRun 
            Caption         =   "&Run"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   12
            Top             =   3240
            Width           =   735
         End
         Begin VB.TextBox txtResults 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   2160
            Width           =   2055
         End
         Begin VB.ComboBox cbo 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.Image img 
            Height          =   300
            Index           =   0
            Left            =   480
            Top             =   960
            Width           =   300
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            Caption         =   "Caption:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1058
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":178C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":265A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2974
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4376
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4710
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5044
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   5
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6585
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ListView:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   3
         Tag             =   " ListView:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Options:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   " TreeView:"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7725
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5953
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "08/15/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "4:17 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2685
      Top             =   3750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Clear"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Sub cmdClear_Click()
    modMain.Setup_Form Me
End Sub

Private Sub cmdRun_Click()
    modMain.Run_WebService Me
End Sub

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    imgSplitter.Left = GetSetting(App.Title, "Settings", "SplitterPos", 3000)
    AddNode Me
    SizeControls imgSplitter.Left
    Me.tvTreeView.Nodes(1).Selected = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    For i = 0 To Me.lblName.UBound
        If Me.lblName(i).FontUnderline Then
            Me.lblName(i).FontUnderline = False
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "SplitterPos", imgSplitter.Left
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub

Private Sub fraMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
    For i = 0 To Me.lblName.UBound
        If Me.lblName(i).FontUnderline Then
            Me.lblName(i).FontUnderline = False
        End If
    Next i
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub

Sub SizeControls(X As Single)
    On Error Resume Next

    'set the width
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    tvTreeView.Width = X
    imgSplitter.Left = X
    picForm.Left = X + 40
    picForm.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = picForm.Left + 20
    lblTitle(1).Width = picForm.Width - 40
        
    
    'set the top
    tvTreeView.Top = picTitles.Height

    picForm.Top = tvTreeView.Top
    
    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    
    picForm.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
    
    Me.fraMain.Height = Me.picForm.Height - 100
    Me.fraMain.Top = 0
    Me.fraMain.Left = 50
    Me.fraMain.Width = Me.picForm.Width - 150
    
    Me.txtResults(0).Width = Me.fraMain.Width - 3000
End Sub

Private Sub lblName_Click(Index As Integer)
Dim StrP As String
Dim iC As Long
Dim MyNode As Node
Dim MyIndex As Integer

    Set MyNode = Me.tvTreeView.SelectedItem
    MyIndex = Index + Me.tvTreeView.SelectedItem.Index + 1
    
    If MyNode.Children > 0 Then
        'StrP = Me.tvTreeView.SelectedItem.Parent.Key
        iC = MyNode.Child.Index
        If MyIndex = iC Then
            Me.tvTreeView.Nodes(iC).Selected = True
            Me.tvTreeView.Nodes(iC).EnsureVisible
            tvTreeView_NodeClick Me.tvTreeView.Nodes(iC)
        Else
            Do While iC <> MyNode.Child.LastSibling.Index
                iC = Me.tvTreeView.Nodes(iC).Next.Index
                If MyIndex = iC Then
                    Me.tvTreeView.Nodes(iC).Selected = True
                    Me.tvTreeView.Nodes(iC).EnsureVisible
                    tvTreeView_NodeClick Me.tvTreeView.Nodes(iC)
                    Exit Do
                End If
            Loop
        End If
    End If
    Set MyNode = Nothing
End Sub

Private Sub lblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.tvTreeView.SelectedItem.Children > 0 Then
        If Not Me.lblName(Index).FontUnderline Then
            Me.lblName(Index).FontUnderline = True
        End If
    Else
        If Me.lblName(Index).FontUnderline Then
            Me.lblName(Index).FontUnderline = False
        End If
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuViewWebBrowser_Click()
    Dim frmB As New frmBrowser
    frmB.StartingAddress = "http://www.webservicex.net/WS/default.aspx"
    frmB.Show
End Sub

Private Sub mnuViewRefresh_Click()
    modMain.Setup_Form Me
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    On Error Resume Next
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    'ToDo: Add 'mnuFileNew_Click' code.
    cmdClear_Click
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: add code to process the opened file

End Sub

Private Sub tvTreeView_GotFocus()
    Call tvTreeView_NodeClick(tvTreeView.SelectedItem)
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    modMain.Setup_Form Me
    
    If Node.Parent Is Nothing Then
        Me.fraMain.Enabled = True
        Me.fraMain.Caption = vbNullString
        Me.lblTitle(1).Caption = Node.Text
    Else
        With Me.fraMain
            .Enabled = True
            .Caption = Node.Text
            .FontBold = True
            .ForeColor = vbBlue
        End With
    End If
End Sub
