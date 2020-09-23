VERSION 5.00
Begin VB.Form frmCaptureArea 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Capture Area"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmCaptureArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'Module:        frmCaptureArea
'Description:   Screen Capture Utility for Quick Spy
'Date:          May 11, 2001
'Last Updated:  May 16, 2001
'Developer:     Shannon Harmon (shannonh@theharmonfamily.com)
'Info:          Feel free to use any of the routines in your
'               own package, but you may not distribute/sell
'               Quick Spy in whole or part without my consent.
'               Copyright 2001, Shannon Harmon - All rights reserved!
'*************************************************************
Option Explicit

Dim ptStart As POINTAPI
Dim ptEnd As POINTAPI
Dim rc As RECT
Dim rc2 As RECT
Dim blnMouseDown As Boolean

Private Sub Form_Load()
  Select Case CA_Cur
    Case CA_CUSTOM
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(102, vbResCursor)
    Case CA_16
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(103, vbResCursor)
    Case CA_32
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(104, vbResCursor)
    Case CA_48
      Me.MousePointer = vbCustom
      Me.MouseIcon = LoadResPicture(105, vbResCursor)
    Case Else
      Unload Me
      Exit Sub
  End Select
  
  Dim dc As Long
  dc = GetWindowDC(GetDesktopWindow())
  BitBlt Me.hdc, 0, 0, Screen.Width, Screen.Height, dc, 0, 0, SRCCOPY
  Me.Picture = Me.Image
  Call ReleaseDC(GetDesktopWindow(), dc)
  Me.AutoRedraw = False
  Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If CA_Cur = CA_CUSTOM Then
    ptStart.x = x
    ptStart.y = y
    blnMouseDown = True
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If CA_Cur = CA_48 Then  'Windows cursor limited to 32x32 so I have to draw this, slower!
    Me.Cls
    Me.ForeColor = vbWhite
    Rectangle Me.hdc, x + 1, y + 1, x + 46, y + 46
    Me.ForeColor = vbBlack
    Rectangle Me.hdc, x, y, x + 48, y + 48
  ElseIf CA_Cur = CA_CUSTOM Then
    If blnMouseDown Then
      Me.Cls
      ptEnd.x = x
      ptEnd.y = y
      
      Me.ForeColor = vbWhite
      Rectangle Me.hdc, ptStart.x + 1, ptStart.y + 1, ptEnd.x - 1, ptEnd.y - 1
      Me.ForeColor = vbBlack
      Rectangle Me.hdc, ptStart.x, ptStart.y, ptEnd.x, ptEnd.y
      
      Call GetRectFromPoints(ptStart, ptEnd, rc)
      If (rc.Bottom - rc.Top) > 100 And (rc.Right - rc.Left) > 100 Then
        rc2 = rc
        rc2.Top = rc.Top + ((rc.Bottom - rc.Top) / 2) - 5
        Call DrawText(Me.hdc, " " & rc.Right - rc.Left & " x " & _
                      rc.Bottom - rc.Top & " " & vbNullChar, -1, rc2, DT_CENTER)
      End If
    End If
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim pic As Picture
  Me.Cls
  
  If CA_Cur <> CA_CUSTOM Then
    ptStart.x = x
    ptStart.y = y
    ptEnd.x = x + (16 * CA_Cur)
    ptEnd.y = y + (16 * CA_Cur)
  Else
    ptEnd.x = x
    ptEnd.y = y
  End If

  Call GetRectFromPoints(ptStart, ptEnd, rc)
  
  If CaptureRect(Me.hdc, rc, pic) = 1 Then
    If frmMain.mnuCapture(5).Checked Then
      Clipboard.Clear
      Clipboard.SetData pic, vbCFBitmap
    Else
      Dim strfile As String
      Me.Hide
      strfile = GetSaveFileName(Me.hwnd)
      
      If strfile <> "" Then
        On Error Resume Next
        DoEvents
        SavePicture pic, strfile
        If Err Then
          MsgBox LoadResString(ERROR_SAVING_IMAGE) & vbCrLf & vbCrLf & _
                 "Number: " & Err.Number & vbCrLf & Err.Description, vbCritical
          Err.Clear
        End If
      End If
    End If
    Set pic = LoadPicture()
  Else
    Me.Hide
    MsgBox LoadResString(ERROR_NO_CAPTURE), vbInformation
  End If
  
  Unload Me
End Sub
