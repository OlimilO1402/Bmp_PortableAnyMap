VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagePNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://de.wikipedia.org/wiki/Portable_Anymap
'PNM = Portable Anymap
'PBM = Portable Bitmap  ' 2 colors black-white  ' P1 asc ' P4 bin
'PGM = Portable Graymap ' 256 indexed greyscale ' P2 asc ' P5 bin
'PPM = Portable Pixmap  ' 8-Bit or 16-Bit color ' P3 asc ' P6 bin
'PNM = Portable Anymap

'P1  Portable Bitmap     ASCII ' Farben: 2
'P2  Portable Graymap    ASCII ' Graustufen: 256
'P3  Portable Pixmap     ASCII ' Farben: 24bit
'P4  Portable Bitmap     Bin‰r
'P5  Portable Graymap    Bin‰r
'P6  Portable Pixmap     Bin‰r
'P7  Portable Anymap     ASCII oder Bin‰r?

Public Enum EPNMType
    PortableBitmapAsc = 1
    PortableGraymapAsc
    PortablePixmapAsc
    PortableBitmapBin
    PortableGraymapBin
    PortablePixmapBin
    PortableAnyMapBin
End Enum

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type TLng
    Value As Long
End Type

Private Type BGRA '
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
Private Type BITMAPINFOHEADER
    biSize          As Long    ' 4
    biWidth         As Long    ' 4
    biHeight        As Long    ' 4
    biPlanes        As Integer ' 2
    biBitCount      As Integer ' 2
    biCompression   As Long    ' 4
    biSizeImage     As Long    ' 4
    biXPelsPerMeter As Long    ' 4
    biYPelsPerMeter As Long    ' 4
    biClrUsed       As Long    ' 4
    biClrImportant  As Long    ' 4
End Type                 ' Sum: 40

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfo
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As BGRA
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/olectl/ns-olectl-pictdesc
Private Type PICTDESC
    cbSize  As Long    ' 4
    picType As Long    ' 4    ' PictureTypeConstants
    hbitmap As LongPtr ' 4, 8 ' HBITMAP
    hpal    As LongPtr ' 4, 8 ' HPALETTE
End Type         ' Sum: 16

#If VBA7 Then
    Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByRef pbmi As BITMAPINFO, ByVal usage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal offset As Long) As Long
    
    Private Declare PtrSafe Function SetDIBits Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal hbitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" ( _
        ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
#Else
    Private Declare Function CreateDIBSection Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByRef pbmi As BITMAPINFO, ByVal usage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal offset As Long) As Long
    
    Private Declare Function SetDIBits Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal hbitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
        ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
#End If
Private m_PFN     As String
Private m_PNMType As EPNMType ' P2: Daten als Text gespeichert, P5: Daten bin‰r gespeichert
Private m_Comment As String
Private m_MaxCols As Long     ' maximale Anzahl Graustufen 0..255 = 8 Bit' oder 0..65535 = 16 Bit
Private m_PicGUID As GUID
Private m_BmpInfo As BITMAPINFO
Private m_PicDesc As PICTDESC
Private m_data()  As Byte

Private Sub Class_Initialize()
    Dim i As Integer
    With m_BmpInfo
        .bmiHeader.biSize = LenB(.bmiHeader)
        .bmiHeader.biPlanes = 1
    End With
    m_MaxCols = 255
    With m_PicDesc
        .cbSize = LenB(m_PicDesc)
        .picType = 1 'PictureTypeConstants.vbPicTypeBitmap
    End With
    With m_PicGUID
        .Data1 = &H7BF80981
        .Data2 = &HBF32:  .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(3) = &HAA: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
    End With
End Sub

Friend Function ToPicture() As StdPicture
Try: On Error GoTo Catch
    With m_PicDesc
        .hbitmap = CreateDIBSection(0&, m_BmpInfo, 0&, 0&, 0&, 0&)
        If .hbitmap = 0 Then Exit Function
        If SetDIBits(0&, .hbitmap, 0&, Me.Height, m_data(0), m_BmpInfo, 0&) = 0 Then Exit Function
    End With
    OleCreatePictureIndirect m_PicDesc, m_PicGUID, 1&, ToPicture
    Exit Function
Catch: MsgBox "Could not convert to Picture"
End Function

'Beispiel f¸r Header:
'P5
'# Created by IrfanView
'800 640
'255
'’“»∂™ùôòòööô
''+)++&
'data startet mit ’=&HD5=213
'und endet mit &=&H26=38

Friend Sub New_(aPFN As String)
    Read aPFN
End Sub

Public Function ToStr() As String
    ToStr = Me.PNMType_ToStr & "; " & m_Comment & "; w: " & Me.Width & "; h: " & Me.Height & "; c: " & m_MaxCols
End Function

Public Sub Read(aPFN As String)
    m_PFN = aPFN
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open m_PFN For Binary Access Read As FNr
    If Not IsAnymap(FNr) Then Exit Sub
    ReadComment FNr 'maybe or maybe not there
    If Not ReadSize(FNr) Then Exit Sub
    If Not ReadMaxCols(FNr) Then Exit Sub
    If Not ReadData(FNr) Then Exit Sub
    GoTo Finally
Catch:
    MsgBox "Error reading file: " & vbCrLf & m_PFN
Finally:
    Close FNr
End Sub

Private Function IsAnymap(FNr As Integer) As Boolean
    Dim ch As Byte
    Get FNr, , ch: IsAnymap = ch = Asc("P")
    If Not IsAnymap Then MsgBox "Not a valid pnm file type.": Exit Function
    Get FNr, , ch:
    m_PNMType = ch - 48:  IsAnymap = m_PNMType < 8
    If Not IsAnymap Then MsgBox "Maybe pnm file but file type not supported: " & m_PNMType
    Get FNr, , ch
    Select Case ch
    Case Asc(vbLf)
    Case Asc(vbCr)
        Get FNr, , ch
        If ch <> Asc(vbLf) Then IsAnymap = False: Exit Function
    End Select
    Dim i As Integer
    With m_BmpInfo
        Select Case m_PNMType
        Case EPNMType.PortableBitmapAsc, EPNMType.PortableBitmapBin
            With .bmiHeader
                .biBitCount = 8
                .biClrUsed = 256
            End With
            .bmiColors(0) = New_BGRA(255, 255, 255, 0) ' Weiﬂ
            .bmiColors(1) = New_BGRA(0, 0, 0, 0)       ' Schwarz
        Case EPNMType.PortableGraymapAsc, EPNMType.PortableGraymapBin
            With .bmiHeader
                .biBitCount = 8
                .biClrUsed = 256
            End With
            For i = 0 To 255 'do not use i as byte -> overflow
                With .bmiColors(i)
                    .R = i: .G = i: .B = i
                End With
            Next
        Case EPNMType.PortablePixmapAsc, EPNMType.PortablePixmapBin
            With .bmiHeader
                .biBitCount = 24
                .biClrUsed = 16777216
            End With
        Case Else
            '???
        End Select
    End With
End Function

Private Function New_BGRA(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal A As Byte) As BGRA
    With New_BGRA
        .R = R: .G = G: .B = B: .A = A
    End With
End Function

Private Sub ReadComment(FNr As Integer)
    Dim sk As Long: sk = Seek(FNr)
    Dim ch As Byte
    Get FNr, , ch
    If ch <> Asc("#") Then
        Get FNr, sk - 1, ch
        Exit Sub
    End If
    Dim iStart As Long, iEnd As Long
    iStart = Seek(FNr)
    Do
        Get FNr, , ch
        Select Case ch
        Case Asc(vbLf)
            iEnd = Seek(FNr)
            m_Comment = Space(iEnd - iStart)
            Get FNr, iStart - 1, m_Comment
            Exit Sub
        Case Asc(vbCr): Get FNr, , ch
            If ch <> Asc(vbLf) Then Exit Sub
        End Select
    Loop Until EOF(FNr)
End Sub

Private Function ReadSize(ByVal FNr As Integer) As Boolean
    Dim W As Long, H As Long
    ReadSize = ReadInteger(FNr, W)
    Me.Width = W
    If ReadSize Then
        ReadSize = ReadInteger(FNr, H)
        Me.Height = H
    End If
End Function

Private Function ReadInteger(ByVal FNr As Integer, ByRef i_out As Long) As Boolean
    Dim ch As Byte, state As Long
    Do
        Get FNr, , ch
        Select Case state
        Case 0 'hat noch nicht begonnen
            Select Case ch
            Case &H20, &HA    ' Weiterlesen
            Case &H30 To &H39 'Zeichen "0" bis "9"
                'jetzt Zahl lesen starten
                i_out = ch - &H30
                state = state + 1
            Case Else 'error
                Exit Do 'Function
            End Select
        Case 1
            'Zahl lesen l‰uft
            Select Case ch
            Case &H20, &HA, &HD
                ReadInteger = True
                'jetzt zuende
                Exit Do 'Function
            Case &H30 To &H39: 'Zeichen "0" bis "9" OK weiterlesen
                i_out = i_out * 10 + (ch - &H30)
            Case Else 'error
                Exit Function
            End Select
        End Select
    Loop Until EOF(FNr)
    ReadInteger = True
End Function

Private Function ReadMaxCols(FNr As Integer) As Boolean
    ReadMaxCols = ReadInteger(FNr, m_MaxCols)
End Function

Private Function ReadData(FNr As Integer) As Boolean
Try: On Error GoTo Catch
    Dim u As Long: u = LOF(FNr) - Seek(FNr)
    If m_PNMType < 4 Then
        'read ascii
        Dim sLines As String: sLines = Space(u + 1)
        Get FNr, , sLines
        Dim Lines() As String
        If InStr(1, sLines, vbCrLf) > 0 Then
            Lines = Split(sLines, vbCrLf)
        ElseIf InStr(1, sLines, vbCr) > 0 Then
            Lines = Split(sLines, vbCr)
        ElseIf InStr(1, sLines, vbLf) > 0 Then
            Lines = Split(sLines, vbLf)
        End If
        Dim sa() As String, s As String
        Dim i As Long, j As Long, c As Long
        ReDim m_data(0 To Me.ByteSize - 1)
        For i = 0 To UBound(Lines)
            s = Lines(i)
            sa = Split(s, " ")
            For j = 0 To UBound(sa)
                s = sa(j)
                If Len(s) And IsNumeric(s) Then
                    m_data(c) = CByte(s)
                    c = c + 1
                End If
            Next
        Next
    Else
        'read binary
        ReDim m_data(0 To u)
        Get FNr, , m_data
    End If
    ReadData = (u + 1) = Me.Width * Me.Height * Me.MaxColors \ 255
    Exit Function
Catch:
    'if err.Number = 9
    MsgBox Err.Number & " " & "Error reading data" & vbCrLf & Err.Description
End Function

'Private Function ReadData(FNr As Integer) As Boolean
'Try: On Error GoTo Catch
'    Dim u As Long: u = LOF(FNr) - Seek(FNr)
'    If m_PNMType < 4 Then
'        'read ascii
'        ReDim m_data(0 To Me.ByteSize)
'        Do Until EOF(FNr)
'            If ReadInteger(FNr, b) Then
'                m_data(i) = CByte(b)
'            End If
'            i = i + 1
'        Loop Until EOF(FNr)
'    Else
'        'read binary
'        ReDim m_data(0 To u)
'        Get FNr, , m_data
'    End If
'    ReadData = (u + 1) = Me.Width * Me.Height * Me.MaxColors \ 255
'    Exit Function
'Catch:
'    'if err.Number = 9
'    MsgBox Err.Number & " " & "Error reading data" & vbCrLf & Err.Description
'End Function

Public Property Get PNMType() As EPNMType
    PNMType = m_PNMType
End Property
Public Function PNMType_ToStr() As String
    Dim s As String
    Select Case m_PNMType
    Case EPNMType.PortableBitmapAsc:  s = "Portable Bitmap (Ascii) P1"
    Case EPNMType.PortableGraymapAsc: s = "Portable Graymap(Ascii) P2"
    Case EPNMType.PortablePixmapAsc:  s = "Portable Bitmap (Ascii) P3"
    Case EPNMType.PortableBitmapBin:  s = "Portable Bitmap(Binary) P4"
    Case EPNMType.PortableGraymapBin: s = "Portable Grayma(Binary) P5"
    Case EPNMType.PortablePixmapBin:  s = "Portable Pixmap(Binary) P6"
    Case EPNMType.PortableAnyMapBin:  s = "Portable AnyMap(Binary) P7"
    End Select
    PNMType_ToStr = s
End Function

Public Property Get ByteSize() As Long
    'ByteSize = UBound(m_data) * IIf(m_MaxCols = 255, 1, 2)
    ByteSize = Me.Width * Me.Height * IIf(m_MaxCols = 255, 1, 2)
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Get Width() As Long
    Width = m_BmpInfo.bmiHeader.biWidth
End Property
Public Property Let Width(ByVal Value As Long)
    m_BmpInfo.bmiHeader.biWidth = Value
End Property

Public Property Get Height() As Long
    Height = -m_BmpInfo.bmiHeader.biHeight
End Property
Public Property Let Height(ByVal Value As Long)
    m_BmpInfo.bmiHeader.biHeight = -Value
End Property

Public Property Get MaxColors() As Long
    MaxColors = m_MaxCols
End Property

'Public Property Get Pixel(ByVal x As Long, ByVal y As Long) As Long
'    Dim c As Byte: c = m_data(y * m_Width + x)
'    Pixel = RGB(c, c, c)
'End Property
'Public Property Let Pixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
'    Dim l As TLng: l.Value = Color
'    Dim c As RGBA: LSet c = l
'    With c
'        m_data(y * m_Width + x) = (.R + .G + .B) \ 3
'    End With
'End Property
