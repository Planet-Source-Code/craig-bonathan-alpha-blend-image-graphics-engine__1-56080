VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABI Graphics Engine"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer RefreshTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   0
      Tag             =   "0"
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Use the '\' and '/' keys to move the DNA. Press 'P' to pause/resume spinning."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   3255
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 2004, Craig Bonathan

Option Explicit

Dim Buffer As New GraphicsBuffer
Dim Images As New Collection
Dim Background As StdPicture

Dim Spinning As Boolean, XPos As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
    Case "\"
        XPos = XPos - 5
        If XPos < -100 Then XPos = 300
    Case "/"
        XPos = XPos + 5
        If XPos > 300 Then XPos = -100
    Case "p"
        Spinning = Not Spinning
    End Select
End Sub

Private Sub Form_Load()
    Dim Temp As AlphaBlendImage, Pos As Long
    
    Main.Show
    DoEvents
    
    ' Load each image from DNA.abi (30 images in total)
    For Pos = 0 To 29
        Set Temp = New AlphaBlendImage
        Temp.LoadImage App.Path & "\DNA.abi", Pos, 150
        Images.Add Temp, "DNA_" & CStr(Pos)
        Set Temp = Nothing
    Next
    
    ' Create graphics buffer with a background
    Set Background = LoadPicture(App.Path & "\DNA.bmp")
    Buffer.CreateBuffer 300, 200, Background.handle
    
    RefreshTimer.Enabled = True
    Spinning = True
    XPos = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Pos As Long
    
    RefreshTimer.Enabled = False
    
    ' Unload the DNA images
    For Pos = 0 To 29
        Images.Item("DNA_" & CStr(Pos)).UnloadImage
    Next
    
    ' Close the graphics buffer
    Buffer.CloseBuffer
    
    Set Background = Nothing
End Sub

Private Sub RefreshTimer_Timer()
    Dim TempDC As Long, Temp As Long
    
    RefreshTimer.Enabled = False
    Temp = CLng(RefreshTimer.Tag)
    
    ' Get Device Context and paste DNA background on to graphics buffer
    TempDC = Buffer.GetBufferDC
    Buffer.EraseToBackground
    
    ' Draw a DNA image on to the graphics buffer
    Images.Item("DNA_" & CStr(Temp)).DrawImage TempDC, XPos, 0
    
    ' Draw the graphics buffer to the form
    Buffer.DrawToDC Main.hDC, 0, 0
    
    ' Note: Going through a buffer to draw images, and then
    ' writing to the display prevents unwanted graphical glitches.
    ' For example, if you clear the display, and then write
    ' something to it, you will see the display flicker.
    
    If Spinning = True Then Temp = (Temp + 1) Mod 30
    RefreshTimer.Tag = CStr(Temp)
    RefreshTimer.Enabled = True
End Sub
