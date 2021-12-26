VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6735
      Left            =   0
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   701
      TabIndex        =   1
      Top             =   600
      Width           =   10575
   End
   Begin VB.CommandButton BtnOpenFolder 
      Caption         =   "Open pgms-subfolder"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Drag'n'drop pictures of filetype pgm, ppm, pbm or pnm into the box."
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_pnm As ImagePNM

Private Sub BtnOpenFolder_Click()
    Dim p As String: p = App.Path & "\pgms\"
    If MsgBox("Open folder?" & vbCrLf & p, vbOKCancel) = vbCancel Then Exit Sub
    Shell "Explorer.exe " & p, vbNormalFocus
End Sub

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim pfn As String: pfn = Data.Files(1)
    Dim dt As Single: dt = Timer
    Set m_pnm = MMain.ImagePNM(pfn)
    dt = Timer - dt
    If m_pnm Is Nothing Then Exit Sub
    Label1.Caption = m_pnm.ToStr & " dt: " & dt & "sec;"
    Set Picture1.Picture = m_pnm.ToPicture
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H
End Sub

'
'Private Type TLng
'    Value As Long
'End Type
'
'Private Type RGBA
'    R As Byte
'    G As Byte
'    b As Byte
'    A As Byte
'End Type
'
'Dim pfn As String
'
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P2_VB-IDE.pgm"              ' OK
'    pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P2_Winterwald.pgm"          ' OK
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P3_PSPFarben.ppm"           ' Problem
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P4_PSPFarben.pbm"           ' convert to Picture problem
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_Grafitty.pgm"            ' OK
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_PSPFarben.pgm"           ' OK
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_sample_640×426.pgm"      ' OK
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_VB-IDE.pgm"              ' OK
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_Winterwald.pgm"          ' OK
'    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P6_PSPFarben.ppm"           ' Mist
'    Dim dt As Single: dt = Timer
'    Dim pic As ImagePNM: Set pic = MMain.ImagePNM(pfn)
'    dt = Timer - dt
'    'MsgBox dt '1,535
'    If pic Is Nothing Then Exit Sub
'    Label1.Caption = pic.ToStr
'    Set Picture1.Picture = pic.ToPicture

'Private Sub Command2_Click()
'    Randomize
'    Dim b As Byte, c1 As TLng, c2 As RGBA
'    Dim i As Long, n As Long: n = 1000
'    For i = 0 To 255
'        b = i 'CByte(Rnd * 255)
'        c1.Value = RGB(b, b, b)
'        LSet c2 = c1
'        If Not RGBA_IsGray(c2, b) Then
'            RGBA_ToStr c2
'        End If
'    Next
'End Sub
'
'Private Function RGBA_IsGray(this As RGBA, ByVal Gray As Byte) As Boolean
'    With this
'        RGBA_IsGray = .R = Gray And .G = Gray And .b = Gray
'    End With
'End Function
'
'Private Function RGBA_ToStr(this As RGBA) As String
'    With this
'        RGBA_ToStr = "RGBA{R: " & .R & "; G: " & .G & "; B: " & .b & "}"
'    End With
'End Function
