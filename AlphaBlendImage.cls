VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaBlendImage"
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

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32.dll" (ByVal hDC As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, ByVal lpInitBits As Long, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Private ImageDC As Long, ImageBitmap As Long, ImageOldBitmap As Long
Private Width As Long, Height As Long
Public ImageName As String
Private BlendOptions As Long
Private TempPic As IPictureDisp
Private Init As Boolean

Public Sub LoadImage(FileName As String, ImageNum As Long, Optional Opacity As Byte = 255)
    Dim ImageCount As Long
    Dim FileNum As Long, DataPos As Long, DataSize As Long
    Dim ImageData() As Byte
    Dim BitmapDetails As BITMAPINFO
    Dim Temp As BLENDFUNCTION
    
    FileNum = FreeFile
    Open FileName For Binary Access Read As #FileNum
    Get #FileNum, 1, ImageCount
    Get #FileNum, 5, Width
    Get #FileNum, 9, Height
    DataSize = Width * Height * 4
    DataPos = DataSize * ImageNum + 12
    ReDim ImageData(DataSize - 1)
    Get #FileNum, DataPos, ImageData()
    Close #FileNum
    
    ImageDC = CreateCompatibleDC(0)
    
    BitmapDetails.bmiHeader.biSize = Len(BitmapDetails.bmiHeader)
    BitmapDetails.bmiHeader.biWidth = Width
    BitmapDetails.bmiHeader.biHeight = Height
    BitmapDetails.bmiHeader.biPlanes = 1
    BitmapDetails.bmiHeader.biBitCount = 32
    BitmapDetails.bmiHeader.biCompression = 0
    BitmapDetails.bmiHeader.biSizeImage = 0
    BitmapDetails.bmiHeader.biXPelsPerMeter = 0
    BitmapDetails.bmiHeader.biYPelsPerMeter = 0
    BitmapDetails.bmiHeader.biClrImportant = 0
    BitmapDetails.bmiHeader.biClrUsed = 0
    
    ImageBitmap = CreateDIBitmap(GetDC(0), BitmapDetails.bmiHeader, &H4, VarPtr(ImageData(0)), BitmapDetails, 0)
    
    If ImageBitmap = 0 Then MsgBox ("Error: Loading Alpha Blend Image")
    ImageOldBitmap = SelectObject(ImageDC, ImageBitmap)
    
    Temp.BlendOp = 0
    Temp.BlendFlags = 0
    Temp.SourceConstantAlpha = Opacity
    Temp.AlphaFormat = 1
    
    CopyMemory VarPtr(BlendOptions), VarPtr(Temp), 4
    
    Init = True
End Sub

Public Sub UnloadImage()
    If Init = True Then
        SelectObject ImageDC, ImageOldBitmap
        DeleteObject ImageBitmap
        DeleteDC ImageDC
        Init = False
    End If
End Sub

Public Sub DrawImage(hDC As Long, XPos As Long, YPos As Long)
    AlphaBlend hDC, XPos, YPos, Width, Height, ImageDC, 0, 0, Width, Height, BlendOptions
End Sub

Private Sub Class_Terminate()
    UnloadImage
End Sub
