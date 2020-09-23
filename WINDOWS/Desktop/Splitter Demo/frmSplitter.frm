VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplitter 
   Caption         =   "Splitter Demo"
   ClientHeight    =   9675
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   6585
   Begin VB.PictureBox picRightPane 
      Align           =   3  'Align Left
      Height          =   9675
      Left            =   1515
      ScaleHeight     =   9615
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   0
      Width           =   1455
      Begin MSComctlLib.ListView lvwMain 
         Height          =   2175
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picSplitter 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   9675
      Left            =   1455
      MousePointer    =   9  'Size W E
      ScaleHeight     =   9675
      ScaleWidth      =   60
      TabIndex        =   1
      Tag             =   "50"
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox picLeftPane 
      Align           =   3  'Align Left
      Height          =   9675
      Left            =   0
      ScaleHeight     =   9615
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin MSComctlLib.TreeView tvwMain 
         Height          =   2775
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   4895
         _Version        =   393217
         Style           =   7
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    Dim nRatio As Integer
    
    If Not Me.WindowState = vbMinimized Then
        'Determine existing ratio of panes
        nRatio = Val(picSplitter.Tag)
        
        picLeftPane.Width = Me.ScaleWidth * nRatio / 100 - 30
        picRightPane.Width = Me.ScaleWidth - picLeftPane.Width - 60
    End If
    
End Sub


Private Sub picLeftPane_Resize()

    'Move treeview (-15 means border is hidden off left/top of pic box)
    tvwMain.Move -15, -15, picLeftPane.ScaleWidth + 30, picLeftPane.ScaleHeight + 30
    
End Sub


Private Sub picRightPane_Resize()

    'Move listview (-15 means border is hidden off left/top of pic box)
    lvwMain.Move -15, -15, picRightPane.ScaleWidth + 30, picRightPane.ScaleHeight + 30
    
End Sub


Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim nNewRatio As Integer
    
    If Button = vbLeftButton Then
        'Determine ratio in new position
        nNewRatio = (picSplitter.Left + X) / Me.ScaleWidth * 100
        
        'Check it is within acceptable bounds
        If nNewRatio > 10 And nNewRatio < 90 Then
            picLeftPane.Width = picSplitter.Left + X
            picRightPane.Width = Me.ScaleWidth - picLeftPane.Width - 60
            picSplitter.Tag = nNewRatio
        End If
    End If
    
End Sub


