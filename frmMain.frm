VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Album Art Tag Editor"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdBefore 
      Caption         =   "<"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtAbout 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1575
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMain.frx":0000
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdDeletePicture 
      Caption         =   "Delete Picture"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddPicture 
      Caption         =   "Add Picture"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   6360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picAlbumArt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   240
      Picture         =   "frmMain.frx":0019
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   3000
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Load MP3"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox cmbPictureType 
      Height          =   315
      ItemData        =   "frmMain.frx":1A37
      Left            =   240
      List            =   "frmMain.frx":1A7A
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblIndex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 / 0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MP3Path    As String
Private TPic       As IPictureDisp
Private CurIndex   As Long
Private MaxIndex   As Long

Private Sub cmdAddPicture_Click()

    cdlMain.Filename = vbNullString
    cdlMain.Filter = "Image File|*.bmp;*.gif;*.jpg"
    cdlMain.ShowOpen
    If LenB(Dir(MP3Path)) > 0 And LenB(cdlMain.Filename) Then
        If Not ID3Exist(MP3Path) Then
            txtAbout.Text = "ID3 v2.3" & vbNewLine & MP3Path
        End If
        Set TPic = LoadPicture(cdlMain.Filename)
        If WriteAlbumArt(MP3Path, CurIndex, TPic, cmbPictureType.ListIndex) Then
            picAlbumArt.Cls
            ResizePic
            MaxIndex = MaxIndex + 1
            CurIndex = CurIndex + 1
            lblIndex = CurIndex & " / " & MaxIndex
        End If
    End If

End Sub

Private Sub cmdBefore_Click()

Dim k As Long

    If CurIndex > 1 Then
        picAlbumArt.Cls
        CurIndex = CurIndex - 1
        lblIndex = CurIndex & " / " & MaxIndex
        If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
            cmbPictureType.ListIndex = k
            ResizePic
        End If
    End If

End Sub

Private Sub cmdDeletePicture_Click()

Dim k As Long

    If LenB(Dir(MP3Path)) Then
        If DeleteAlbumArt(MP3Path, CurIndex) Then
            picAlbumArt.Cls
            MaxIndex = MaxIndex - 1
            If CurIndex - 1 > 0 Then
                CurIndex = CurIndex - 1
                If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                    cmbPictureType.ListIndex = k
                    ResizePic
                End If
                lblIndex = CurIndex & " / " & MaxIndex
            ElseIf MaxIndex > 0 Then
                If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                    cmbPictureType.ListIndex = k
                    ResizePic
                End If
            Else
                CurIndex = 0
            End If
            lblIndex = CurIndex & " / " & MaxIndex
        End If
    End If

End Sub

Private Sub cmdNext_Click()

Dim k As Long

    If CurIndex < MaxIndex Then
        picAlbumArt.Cls
        CurIndex = CurIndex + 1
        lblIndex = CurIndex & " / " & MaxIndex
        If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
            cmbPictureType.ListIndex = k
            ResizePic
        End If
    End If

End Sub

Private Sub cmdOpen_Click()

Dim k As Long

    On Error Resume Next
    cdlMain.Filename = vbNullString
    cdlMain.Filter = "MP3 File|*.mp3"
    cdlMain.ShowOpen
    If LenB(cdlMain.Filename) Then
        MP3Path = cdlMain.Filename
        picAlbumArt.Cls
        If ID3Exist(MP3Path) Then
            MaxIndex = GetAlbumArtCount(MP3Path)
            If MaxIndex > 0 Then
                CurIndex = 1
                If ReadAlbumArt(MP3Path, CurIndex, TPic, k) Then
                    cmbPictureType.ListIndex = k
                    ResizePic
                End If
            Else
                CurIndex = 0
            End If
            txtAbout.Text = "ID3 v2." & TVersion & vbNewLine & MP3Path
        Else
            CurIndex = 0
            MaxIndex = 0
            txtAbout.Text = "No ID3v2 Tag" & vbNewLine & MP3Path
        End If
        lblIndex = CurIndex & " / " & MaxIndex
        cdlMain.InitDir = vbNullString
        cmdAddPicture.Enabled = True
        cmdDeletePicture.Enabled = True
    End If
    On Error GoTo 0

End Sub

Private Sub Form_Load()

    cdlMain.flags = &H1000&
    cmbPictureType.ListIndex = 0
    cdlMain.InitDir = App.Path & "\"

End Sub

Private Sub ResizePic()

Dim nWidth  As Long
Dim nHeight As Long

    On Error Resume Next
    nWidth = ScaleX(TPic.Width, vbHimetric, vbPixels)
    nHeight = ScaleY(TPic.Height, vbHimetric, vbPixels)
    With picAlbumArt
        If .ScaleWidth < (nWidth * (.ScaleHeight / nHeight)) Then
            nHeight = nHeight * (.ScaleWidth / nWidth)
            nWidth = .ScaleWidth
        Else
            nWidth = nWidth * (.ScaleHeight / nHeight)
            nHeight = .ScaleHeight
        End If
        TPic.Render .hDC, (.ScaleWidth - CLng(nWidth)) / 2, (.ScaleHeight - CLng(nHeight)) / 2, CLng(nWidth), CLng(nHeight), 0, TPic.Height, TPic.Width, -TPic.Height, ByVal 0&
    End With
    On Error GoTo 0

End Sub
