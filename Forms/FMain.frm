VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "PortableAnyMap"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
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
      BackColor       =   &H00000040&
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

'Private Sub Command1_Click()
'    MsgBox "Das erste und das letzte Byte sind: " & vbCrLf & _
'            "&H" & Hex(m_pnm.Data(0, 0)) & " " & "&H" & Hex(m_pnm.Data(m_pnm.Width - 1, m_pnm.Height - 1))
'End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim pfn As String: pfn = Data.Files(1)
    Dim dt As Single: dt = Timer
    Set m_pnm = MMain.ImagePNM(pfn)
    dt = Timer - dt
    If m_pnm Is Nothing Then Exit Sub
    Set Picture1.Picture = m_pnm.ToPicture
    Label1.Caption = m_pnm.ToStr & " dt: " & dt & "sec;"
    Me.Caption = "PortableAnyMap - " & pfn
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H
End Sub
