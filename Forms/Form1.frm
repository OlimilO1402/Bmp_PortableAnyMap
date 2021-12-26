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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6615
      Left            =   120
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   685
      TabIndex        =   1
      Top             =   600
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TLng
    Value As Long
End Type

Private Type RGBA
    R As Byte
    G As Byte
    b As Byte
    A As Byte
End Type

Dim pfn As String

Private Sub Command1_Click()
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P2_VB-IDE.pgm"              ' OK
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P2_Winterwald.pgm"          ' OK
    pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P3_PSPFarben.ppm"           ' Problem
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P4_PSPFarben.pbm"           ' convert to Picture problem
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_Grafitty.pgm"            ' OK
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_PSPFarben.pgm"           ' OK
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_sample_640×426.pgm"      ' OK
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_VB-IDE.pgm"              ' OK
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P5_Winterwald.pgm"          ' OK
    'pfn = "\\SOLS_DS\Daten\GitHubRepos\VB\Bmp_SIFT\pgms\P6_PSPFarben.ppm"           ' Mist
    
    Dim pic As ImagePNM: Set pic = MMain.ImagePNM(pfn)
    If pic Is Nothing Then Exit Sub
    Command1.Caption = pic.ToStr
    Set Picture1.Picture = pic.ToPicture
    
End Sub

Private Sub Command2_Click()
    Randomize
    Dim b As Byte, c1 As TLng, c2 As RGBA
    Dim i As Long, n As Long: n = 1000
    For i = 0 To 255
        b = i 'CByte(Rnd * 255)
        c1.Value = RGB(b, b, b)
        LSet c2 = c1
        If Not RGBA_IsGray(c2, b) Then
            RGBA_ToStr c2
        End If
    Next
End Sub
Private Function RGBA_IsGray(this As RGBA, ByVal Gray As Byte) As Boolean
    With this
        RGBA_IsGray = .R = Gray And .G = Gray And .b = Gray
    End With
End Function
Private Function RGBA_ToStr(this As RGBA) As String
    With this
        RGBA_ToStr = "RGBA{R: " & .R & "; G: " & .G & "; B: " & .b & "}"
    End With
End Function
    
