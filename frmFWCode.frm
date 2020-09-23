VERSION 5.00
Begin VB.Form frmFWCode 
   Caption         =   "Find Window VB Code Output"
   ClientHeight    =   5265
   ClientLeft      =   1005
   ClientTop       =   1275
   ClientWidth     =   7095
   Icon            =   "frmFWCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7095
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFW 
      Height          =   2940
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   105
      Width           =   4470
   End
End
Attribute VB_Name = "frmFWCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'Module:        frmFWCode
'Description:   Find Window code generator Utility for Quick Spy
'Date:          May 11, 2001
'Last Updated:  May 16, 2001
'Developer:     Shannon Harmon (shannonh@theharmonfamily.com)
'Info:          Feel free to use any of the routines in your
'               own package, but you may not distribute/sell
'               Quick Spy in whole or part without my consent.
'               Copyright 2001, Shannon Harmon - All rights reserved!
'*************************************************************
Option Explicit

Private Sub Form_Load()
  Dim i As Integer
  Dim lhWnd As Long
  Dim lParent As Long
  Dim aWindows() As Long
  Dim strTemp As String
  Dim strClass As String
  Dim strCaption As String
  Dim lPos As Long
  
  If IsNumeric(frmMain.txtHwndLng) Then
    lhWnd = CLng(frmMain.txtHwndLng)
  Else
    lhWnd = 0
  End If
  
  If IsWindow(lhWnd) = 0 Then
    txtFW.Text = LoadResString(ERROR_INVALID_WINDOW)
  Else
    ReDim aWindows(0)
    aWindows(0) = lhWnd
    lParent = GetParent(lhWnd)
    
    Do While lParent <> 0
      ReDim Preserve aWindows(UBound(aWindows) + 1)
      aWindows(UBound(aWindows)) = lParent
      lParent = GetParent(lParent)
    Loop
    
    strTemp = "Option Explicit" & vbNewLine & vbNewLine & "Private Declare Function "
    strTemp = strTemp & "FindWindow Lib ""user32"" Alias ""FindWindowA"" "
    strTemp = strTemp & "(ByVal lpClassName As String, "
    strTemp = strTemp & "ByVal lpWindowName As String) As Long" & vbNewLine
    strTemp = strTemp & "Private Declare Function FindWindowEx "
    strTemp = strTemp & "Lib ""user32"" Alias ""FindWindowExA"" "
    strTemp = strTemp & "(ByVal hWnd1 As Long, ByVal hWnd2 As Long, "
    strTemp = strTemp & "ByVal lpsz1 As String, ByVal lpsz2 As String) As Long"
    strTemp = strTemp & vbNewLine & vbNewLine
    
    strTemp = strTemp & "Public Function FindWindowHwnd() As Long" & vbNewLine
    strTemp = strTemp & vbTab & "Dim lRet As Long" & vbNewLine & vbNewLine
    
    For i = UBound(aWindows) To LBound(aWindows) Step -1
      strClass = Space(256)
      Call GetClassName(aWindows(i), strClass, 256)
      strCaption = String$(255, vbNullChar)
      Call GetWindowText(aWindows(i), strCaption, 255)
      
      lPos = InStr(1, strClass, vbNullChar)
      If lPos > 0 Then strClass = Mid(strClass, 1, lPos - 1)
      lPos = InStr(1, strCaption, vbNullChar)
      If lPos > 0 Then strCaption = Mid(strCaption, 1, lPos - 1)
      
      If i = UBound(aWindows) Then
        strTemp = strTemp & vbTab & "lRet = FindWindow(""" & strClass
        strTemp = strTemp & """, """ & strCaption & """)" & vbNewLine
      Else
        strTemp = strTemp & vbTab & "lRet = FindWindowEx(lRet, 0&,"
        strTemp = strTemp & """" & strClass & """, """ & strCaption
        strTemp = strTemp & """)" & vbNewLine
      End If
    Next i
    
    strTemp = strTemp & vbTab & "FindWindowHwnd = lRet" & vbNewLine
    strTemp = strTemp & "End Function"
    strTemp = strTemp & vbNewLine & vbNewLine
    strTemp = strTemp & "'Notice: This code may or may not work depending on what "
    strTemp = strTemp & "window you use it with." & vbNewLine
    strTemp = strTemp & "'Some windows do not retain their "
    strTemp = strTemp & "same captions 100% of the time, and some are even changed "
    strTemp = strTemp & "by the application itself."
    strTemp = strTemp & vbNewLine & "'You can replace the caption "
    strTemp = strTemp & "with vbNullString in some cases, in others you cannot!"
    strTemp = strTemp & vbNewLine & "'This is by all means only a template to help make it easier!"
    
    txtFW.Text = strTemp
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    txtFW.Move 30, 30, Me.ScaleWidth - 60, Me.ScaleHeight - 60
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Visible = True
  If frmMain.mnuFile(2).Checked Then frmZoom.Visible = True
End Sub
