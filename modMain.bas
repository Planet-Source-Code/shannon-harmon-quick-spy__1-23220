Attribute VB_Name = "modMain"
'*************************************************************
'Module:        modMain
'Description:   Public functions/variables for Quick Spy
'Date:          May 11, 2001
'Last Updated:  May 16, 2001
'Developer:     Shannon Harmon (shannonh@theharmonfamily.com)
'Info:          Feel free to use any of the routines in your
'               own package, but you may not distribute/sell
'               Quick Spy in whole or part without my consent.
'               Copyright 2001, Shannon Harmon - All rights reserved!
'*************************************************************
Option Explicit

Public Const PS_SOLID = 0
Public Const BS_HOLLOW = 1
Public Const HS_SOLID = 8
Public Const RO_COPYPEN = 13
Public Const DT_CENTER = &H1

Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_GETPASSWORDCHAR = &HD2

Public Const ERROR_INVALID_WINDOW = 101
Public Const ERROR_NO_WINDOW_SELECTED = 102
Public Const ERROR_INVALID_COLOR = 103
Public Const ERROR_NO_CAPTURE = 104
Public Const ERROR_SAVING_IMAGE = 105
Public Const ERROR_CANNOT_USE_WITHIN_APP = 106

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_INVALIDATE = &H1
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_NOERASE = &H20
Public Const RDW_NOFRAME = &H800
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_UPDATENOW = &H100
Public Const RDW_VALIDATE = &H8
Public Const RDW_FLAGS = RDW_ALLCHILDREN Or RDW_ERASENOW Or _
             RDW_INTERNALPAINT Or RDW_INVALIDATE Or _
             RDW_FRAME Or RDW_UPDATENOW

Public Const SRCCOPY = &HCC0020

Public Const SWP_NOSIZE = 1
Public Const SWP_NOMOVE = 2

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Public Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(255) As PALETTEENTRY
End Type

Public Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Type PictureBMP
  Size As Long
  Type As Long
  lnghBMP As Long
  lnghPal As Long
  Reserved As Long
End Type

Public Type LOGBRUSH
  lbStyle As Long
  lbColor As Long
  lbHatch As Long
End Type

Public Enum CaptureArea
  CA_UNKNOWN = 0
  CA_16 = 1
  CA_32 = 2
  CA_48 = 3
  CA_CUSTOM = 4
End Enum

Global CA_Cur As CaptureArea
Global picZoomSave As Picture

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PictureBMP, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal lnghPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StringFromGUID2 Lib "ole32" (ByRef lpGUID As GUID, ByVal lpStr As String, ByVal lSize As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Function GetSaveFileName(ByVal hwndOwner As Long) As String
  Dim FileDialog As clsFileDialog
  Set FileDialog = New clsFileDialog
  Static strInitDir As String
  
  With FileDialog
    If strInitDir <> "" Then .InitialDir = strInitDir
    .DefaultExt = "bmp"
    .DialogTitle = App.Title & " - Save Captured Image"
    .Filter = "Bitmap files (*.bmp)|*.bmp"
    .FilterIndex = 0
    .Flags = FlePathMustExist + FleOverWritePrompt
    .hWndParent = hwndOwner
    .MaxFileSize = 255
    If .Show(False) Then
      strInitDir = .FileName
      GetSaveFileName = .FileName
    Else
      GetSaveFileName = ""
    End If
  End With
  
  Set FileDialog = Nothing
End Function

Public Function GetGUID() As String
  Dim udtGUID As GUID
  Dim strGUID As String * 80
  Dim lRet As Long
  
  If CoCreateGuid(udtGUID) = 0 Then
    lRet = StringFromGUID2(udtGUID, strGUID, 80)
    If lRet <> 0 Then
      strGUID = StrConv(strGUID, vbFromUnicode)
      GetGUID = Mid(strGUID, 1, lRet - 1)
    End If
  End If
End Function

Public Function GetRectFromPoints(pt1 As POINTAPI, pt2 As POINTAPI, rc As RECT) As Long
  On Error GoTo PROC_ERR
  
  rc.Left = IIf(pt1.x < pt2.x, pt1.x, pt2.x)
  rc.Top = IIf(pt1.y < pt2.y, pt1.y, pt2.y)
  rc.Right = IIf(pt1.x > pt2.x, pt1.x, pt2.x)
  rc.Bottom = IIf(pt1.y > pt2.y, pt1.y, pt2.y)
  GetRectFromPoints = 1
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  GetRectFromPoints = 0
  Resume PROC_EXIT
End Function

Public Function DrawRect(dc As Long, rc As RECT, Optional ByVal intWidth As Integer = 1, Optional ByVal lngColor As OLE_COLOR = 0&) As Long
  On Error GoTo PROC_ERR
  Dim hPen As Long
  Dim hPenOld As Long
  Dim hBrush As Long
  Dim hBrushOld As Long
  Dim hBrushPrev As Long
  Dim lngROP2ModeOld As Long
  Dim lngResult As Long
  Dim lb As LOGBRUSH
  
  If dc = 0 Then GoTo PROC_EXIT
  
  hPen = CreatePen(PS_SOLID, intWidth, lngColor)
  hPenOld = SelectObject(dc, hPen)
  If hPenOld = 0 Then GoTo PROC_EXIT
  
  lb.lbHatch = HS_SOLID
  lb.lbStyle = BS_HOLLOW
  lb.lbColor = 0& 'Not using, hollow rectangle:)
  
  If GetROP2(dc) <> RO_COPYPEN Then
    lngROP2ModeOld = SetROP2(dc, RO_COPYPEN)
  End If
  
  hBrush = CreateBrushIndirect(lb)
  hBrushOld = SelectObject(dc, hBrush)
  lngResult = Rectangle(dc, rc.Left, rc.Top, rc.Right, rc.Bottom)
  
  DrawRect = 1
    
PROC_EXIT:
  If lngROP2ModeOld <> 0 Then SetROP2 dc, lngROP2ModeOld
  If hBrushOld <> 0 Then SelectObject dc, hBrushOld
  If hPenOld <> 0 Then SelectObject dc, hPenOld
  If hBrush <> 0 Then DeleteObject hBrush
  If hPen <> 0 Then DeleteObject hPen
  Exit Function

PROC_ERR:
  DrawRect = 0
  Resume PROC_EXIT
End Function

Public Function ResizeRect(rc As RECT, ByVal iPixels As Integer) As Long
  On Error GoTo PROC_ERR

  With rc
    .Left = .Left - iPixels
    .Top = .Top - iPixels
    .Right = .Right + iPixels
    .Bottom = .Bottom + iPixels
  End With

  ResizeRect = 1

PROC_EXIT:
  Exit Function

PROC_ERR:
  ResizeRect = 0
  Resume PROC_EXIT
End Function

Public Sub Pause(ByVal lMillSecs As Long)
  On Error GoTo PROC_ERR
  Dim lTime As Long
  lTime = GetTickCount()

  Do While GetTickCount() - lTime < lMillSecs
    DoEvents
  Loop

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Resume PROC_EXIT
End Sub

Public Function CaptureRect(ByVal dc As Long, rc As RECT, pic As Picture) As Long
  On Error GoTo PROC_ERR
  Dim lnghdcMem As Long
  Dim lnghBMP As Long
  Dim lnghBMPPrev As Long
  Dim lngRetval As Long
  Dim lnghPal As Long
  Dim lnghPalPrev As Long
  Dim lngRasterCapsScreen As Long
  Dim lngPaletteScreen As Long
  Dim lngPaletteSizeScreen As Long
  Dim LogPal As LOGPALETTE
  Const clngVGASize As Long = 256

  lngRasterCapsScreen = GetDeviceCaps(dc, RASTERCAPS)
  lngPaletteScreen = lngRasterCapsScreen And RC_PALETTE
  lngPaletteSizeScreen = GetDeviceCaps(dc, SIZEPALETTE)
  lnghdcMem = CreateCompatibleDC(dc)
  lnghBMP = CreateCompatibleBitmap(dc, rc.Right - rc.Left, rc.Bottom - rc.Top)
  lnghBMPPrev = SelectObject(lnghdcMem, lnghBMP)

  If lngPaletteScreen Then
    If (lngPaletteSizeScreen = clngVGASize) Then
      LogPal.palVersion = &H300
      LogPal.palNumEntries = clngVGASize
      lngRetval = GetSystemPaletteEntries(dc, 0, clngVGASize, LogPal.palPalEntry(0))
      lnghPal = CreatePalette(LogPal)
      lnghPalPrev = SelectPalette(lnghdcMem, lnghPal, 0)
      lngRetval = RealizePalette(lnghdcMem)
    End If
  End If

  lngRetval = BitBlt(lnghdcMem, 0, 0, rc.Right - rc.Left, _
              rc.Bottom - rc.Top, dc, rc.Left, rc.Top, vbSrcCopy)
  
  lnghBMP = SelectObject(lnghdcMem, lnghBMPPrev)

  If lngPaletteScreen Then
    If (lngPaletteSizeScreen = clngVGASize) Then
      lnghPal = SelectPalette(lnghdcMem, lnghPalPrev, 0)
    End If
  End If

  lngRetval = DeleteDC(lnghdcMem)
  CaptureRect = CreatePictureFromBitmap(lnghBMP, lnghPal, pic)
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  Set pic = LoadPicture()
  CaptureRect = 0
End Function

Public Function CreatePictureFromBitmap(ByVal lnghBMP As Long, ByVal lnghPal As Long, pic As Picture) As Long
  On Error GoTo PROC_ERR
  Dim lngRetval As Long
  Dim picBMP As PictureBMP
  Dim IPic As IPicture
  Dim IID_IDispatch As GUID

  With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
  End With

  With picBMP
    .Size = Len(picBMP)
    .Type = vbPicTypeBitmap
    .lnghBMP = lnghBMP
    .lnghPal = lnghPal
  End With

  lngRetval = OleCreatePictureIndirect(picBMP, IID_IDispatch, 1, IPic)
  Set pic = IPic
  CreatePictureFromBitmap = 1
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Set pic = LoadPicture()
  CreatePictureFromBitmap = 0
End Function

Public Function WebColor(ByVal lngColor As Long) As String
  If lngColor < vbBlack Or lngColor > vbWhite Then
    Err.Raise vbObjectError + ERROR_INVALID_COLOR, "WebColor", _
              LoadResString(ERROR_INVALID_COLOR)
  Else
    WebColor = Format$(Hex(RGBRed(lngColor)), "00") & _
               Format$(Hex(RGBGreen(lngColor)), "00") & _
               Format$(Hex(RGBBlue(lngColor)), "00")
  End If
End Function

Public Function RGBRed(ByVal lngColor As Long) As Integer
  If lngColor < vbBlack Or lngColor > vbWhite Then
    Err.Raise vbObjectError + ERROR_INVALID_COLOR, "RGBRed", _
              LoadResString(ERROR_INVALID_COLOR)
  Else
    RGBRed = lngColor And &HFF
  End If
End Function

Public Function RGBGreen(ByVal lngColor As Long) As Integer
  If lngColor < vbBlack Or lngColor > vbWhite Then
    Err.Raise vbObjectError + ERROR_INVALID_COLOR, "RGBGreen", _
              LoadResString(ERROR_INVALID_COLOR)
  Else
    RGBGreen = ((lngColor And &H100FF00) / &H100)
  End If
End Function

Public Function RGBBlue(ByVal lngColor As Long) As Integer
  If lngColor < vbBlack Or lngColor > vbWhite Then
    Err.Raise vbObjectError + ERROR_INVALID_COLOR, "RGBBlue", _
              LoadResString(ERROR_INVALID_COLOR)
  Else
    RGBBlue = (lngColor And &HFF0000) / &H10000
  End If
End Function

Public Function GetVBColorName(ByVal lngColor As Long, Optional ByVal strNA As String) As String
  Dim strCol As String
  
  Select Case lngColor
    Case vbBlack:                   strCol = "vbBlack"
    Case vbRed:                     strCol = "vbRed"
    Case vbGreen:                   strCol = "vbGreen"
    Case vbYellow:                  strCol = "vbYellow"
    Case vbBlue:                    strCol = "vbBlue"
    Case vbMagenta:                 strCol = "vbMagenta"
    Case vbCyan:                    strCol = "vbCyan"
    Case vbWhite:                   strCol = "vbWhite"
    Case Else:                      strCol = strNA
  End Select
  
  GetVBColorName = strCol
End Function
