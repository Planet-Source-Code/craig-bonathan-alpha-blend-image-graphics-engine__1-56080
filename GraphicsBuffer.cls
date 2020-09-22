VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2004, Craig Bonathan

Option Explicit

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Init As Boolean
Private BufferDC As Long, BufferBitmap As Long, BufferOldBitmap As Long
Private BackgroundDC As Long, BackgroundBitmap As Long, BackgroundOldBitmap As Long, BackgroundInit As Boolean
Private BufferWidth As Long, BufferHeight As Long


Public Sub CreateBuffer(Width As Long, Height As Long, Optional Background As Long = 0)
    BufferDC = CreateCompatibleDC(0)
    BufferBitmap = CreateCustomBitmap(Width, Height, 24)
    BufferOldBitmap = SelectObject(BufferDC, BufferBitmap)
    BufferWidth = Width
    BufferHeight = Height
    Init = True
    
    If Background <> 0 Then
        BackgroundDC = CreateCompatibleDC(0)
        BackgroundBitmap = Background
        BackgroundOldBitmap = SelectObject(BackgroundDC, BackgroundBitmap)
        BackgroundInit = True
    End If
End Sub

Public Sub CloseBuffer()
    If Init = True Then
        SelectObject BufferDC, BufferOldBitmap
        DeleteObject BufferBitmap
        DeleteDC BufferDC
        Init = False
        
        If BackgroundInit = True Then
            SelectObject BackgroundDC, BackgroundOldBitmap
            DeleteDC BackgroundDC
            BackgroundInit = False
        End If
    End If
End Sub

Private Sub Class_Terminate()
    CloseBuffer
End Sub

Public Function GetBufferDC() As Long
    If Init = True Then GetBufferDC = BufferDC
End Function

Public Sub EraseToBackground()
    If BackgroundInit = True Then BitBlt BufferDC, 0, 0, BufferWidth, BufferHeight, BackgroundDC, 0, 0, vbSrcCopy
End Sub

Public Sub DrawToDC(DestinationDC As Long, X As Long, Y As Long)
    If Init = True Then BitBlt DestinationDC, X, Y, BufferWidth, BufferHeight, BufferDC, 0, 0, vbSrcCopy
End Sub

Private Function CreateCustomBitmap(Width As Long, Height As Long, ColourDepth As Long) As Long
    Dim BitmapDetails As BITMAPINFO
    BitmapDetails.bmiHeader.biSize = Len(BitmapDetails.bmiHeader)
    BitmapDetails.bmiHeader.biBitCount = ColourDepth
    BitmapDetails.bmiHeader.biHeight = Height
    BitmapDetails.bmiHeader.biPlanes = 1
    BitmapDetails.bmiHeader.biWidth = Width
    CreateCustomBitmap = CreateDIBSection(0, BitmapDetails, 0, 0, 0, 0)
End Function
