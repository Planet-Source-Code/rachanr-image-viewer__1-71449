VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Viewer"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5445
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Image Path"
            TextSave        =   "Image Path"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "Dimensions"
            TextSave        =   "Dimensions"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2205
            MinWidth        =   2205
            Text            =   "File Size"
            TextSave        =   "File Size"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview Area "
      Height          =   5295
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   5175
      Begin VB.Image imgPreview 
         BorderStyle     =   1  'Fixed Single
         Height          =   4935
         Left            =   120
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.FileListBox lsbFile 
      Height          =   3210
      Left            =   120
      Pattern         =   "*.bmp;*.jpg"
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.DirListBox lsbDir 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox lsbDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image imgTmp 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   2400
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim picPath As String, n As Long, i As Long
Dim rWidth As Long, rHeight As Long
Dim Ratio As Single, xRatio As Single, yRatio As Single
Dim tmpStr As String, scrX As Long, scrY As Long

Private Sub Form_Load()
scrX = Screen.TwipsPerPixelX
scrY = Screen.TwipsPerPixelY
lsbDir.Path = "C:\"
End Sub

Private Sub lsbDir_Change()
On Error Resume Next
lsbFile.Path = lsbDir.Path
End Sub

Private Sub lsbDrive_Change()
On Error Resume Next
lsbDir.Path = lsbDrive.Drive
End Sub

Private Sub lsbFile_Click()
If Right$(lsbDir.Path, 1) = "\" Then
    picPath = lsbDir.Path & lsbFile.FileName
Else
    picPath = lsbDir.Path & "\" & lsbFile.FileName
End If
imgTmp.Picture = LoadPicture(picPath)
rWidth = imgTmp.Width
rHeight = imgTmp.Height
Ratio = rWidth / rHeight
stbStatus.Panels(1).Text = picPath
stbStatus.Panels(1).ToolTipText = lsbFile.FileName
stbStatus.Panels(2).Text = rWidth / scrX - 4 & " x " & rHeight / scrY - 4
With imgPreview
If Ratio >= 1 Then
    If rWidth < 4935 Then
        xRatio = rWidth
    Else
        xRatio = 4935
    End If
    yRatio = xRatio / Ratio
    .Width = xRatio
    .Height = yRatio
    .Stretch = True
Else
    If rWidth < 4935 Then
        yRatio = rWidth
    Else
        yRatio = 4935
    End If
    xRatio = Ratio * yRatio
    .Width = xRatio
    .Height = yRatio
    .Stretch = True
End If
.Left = Frame1.Width / 2 - xRatio / 2
.Top = 60 + Frame1.Height / 2 - yRatio / 2
.Picture = imgTmp.Picture
End With
End Sub
