VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Background Remover      -      zaidmarkabi@yahoo.com      -      yazanmarkabi.com      -      Arabic Syrian"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox OriginalPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6720
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   1
      ToolTipText     =   "Temporary Image"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remove Setting"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   9015
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1200
         Max             =   200
         TabIndex        =   8
         Top             =   360
         Value           =   40
         Width           =   2175
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   1200
         Max             =   100
         Min             =   10
         TabIndex        =   7
         Top             =   720
         Value           =   40
         Width           =   2175
      End
      Begin VB.CheckBox ChkPrev 
         Caption         =   "Preview Mode"
         Height          =   255
         Left            =   7200
         TabIndex        =   6
         ToolTipText     =   "Preview Mode : Faster"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remove >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "RUN   !"
         Top             =   240
         Width           =   2295
      End
      Begin VB.PictureBox ProgressBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DrawWidth       =   15
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   3960
         ScaleHeight     =   3
         ScaleMode       =   0  'User
         ScaleWidth      =   4905
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "Threshold :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblThre 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Detector :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label LblDet 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Copy to Clipboard"
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   14
      ToolTipText     =   "the Original Image"
      Top             =   480
      Width           =   9000
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   3180
      Left            =   9240
      Picture         =   "frmMain.frx":B03CA
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   13
      ToolTipText     =   "This part should cover Large Area of background As Possible"
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox PreviewPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9240
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   2
      ToolTipText     =   "this image for Preview Mode (tick box ; Faster)"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   " Original Image"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   " Removed Part"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   16
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404040&
      Caption         =   " Preview"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      Top             =   4560
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Image Background Remover - 2010.9.20
'
' This code let you to Remove Background of Image by selecting a Part of This Background.
' then the application will remove the Whole Background automatically.
' Very Easy tool, Fast Simple ans Powerful to Extract Objects from images.
'
' written by  Zaid Markabi
'  Arabic Syrian student
'
' Email : zaidmarkabi@yahoo.com  or  zaid@yazanmarkabi.com
' Website : yazanmarkabi.com  or  yazanmarkabi.webs.com
'

Dim ImageData() As Byte
Dim ImageDataRem() As Byte

Function ColorMustRemoved(ColorR As Byte, ColorG As Byte, ColorB As Byte, Threshold As Byte) As Boolean
Dim QuickX As Long
Dim R As Long, G As Long, B As Long
    For x = 0 To Picture2.ScaleWidth - 1 Step HScroll2.Value
        QuickX = x * 3
    For y = 0 To Picture2.ScaleHeight - 1 Step HScroll2.Value
        R = ImageDataRem(QuickX + 2, y)
        G = ImageDataRem(QuickX + 1, y)
        B = ImageDataRem(QuickX, y)
        If Sqr((ColorR - R) ^ 2) < Threshold And Sqr((ColorG - G) ^ 2) < Threshold And Sqr((ColorB - B) ^ 2) < Threshold Then
            ColorMustRemoved = True
            Exit Function
        End If
    Next
Next
End Function

Private Sub Command1_Click()
Dim x As Long, y As Long
Dim iWidth As Single, iHeight As Single

Picture1.Picture = OriginalPicture.Picture

    PreviewPic.Width = Picture1.Width * 0.25
    PreviewPic.Height = Picture1.Height * 0.25
PreviewPic.PaintPicture Picture1.Picture, 0, 0, PreviewPic.ScaleWidth, PreviewPic.ScaleHeight

If ChkPrev.Value = 1 Then
    PreviewPic.Picture = PreviewPic.Image
    GetImageData2D PreviewPic, ImageData()
    iWidth = PreviewPic.ScaleWidth
    iHeight = PreviewPic.ScaleHeight
Else
    GetImageData2D Picture1, ImageData()
    iWidth = Picture1.ScaleWidth
    iHeight = Picture1.ScaleHeight
End If

GetImageData2D Picture2, ImageDataRem()

Dim R As Byte, G As Byte, B As Byte

ProgressBar.ScaleWidth = iWidth

Dim QuickX As Long

    For x = 0 To iWidth - 1
        QuickX = x * 3
    For y = 0 To iHeight - 1
        R = ImageData(QuickX + 2, y)
        G = ImageData(QuickX + 1, y)
        B = ImageData(QuickX, y)
        If ColorMustRemoved(R, G, B, HScroll1.Value) = True Then
            R = 0: G = 0: B = 0
        End If
        ImageData(QuickX + 2, y) = R
        ImageData(QuickX + 1, y) = G
        ImageData(QuickX, y) = B
    Next y
        DoEvents
        ProgressBar.Line (0, 1)-(x, 1)
    Next x

If ChkPrev.Value = 1 Then
    SetImageData2D PreviewPic, PreviewPic.ScaleWidth, PreviewPic.ScaleHeight, ImageData()
Else
    SetImageData2D Picture1, Picture1.ScaleWidth, Picture1.ScaleHeight, ImageData()
End If

ProgressBar.Cls
End Sub

Private Sub Command4_Click()
Clipboard.Clear
Clipboard.SetData Picture1.Picture
End Sub

Private Sub Form_Load()
OriginalPicture.Picture = Picture1.Picture

Command1_Click
ChkPrev.Value = 0
End Sub

Private Sub HScroll1_Change()
LblThre.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub HScroll2_Change()
LblDet.Caption = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

