VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Spy"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   Begin VB.Frame Frame2 
      Caption         =   "Window Information"
      Height          =   3285
      Left            =   60
      TabIndex        =   4
      Top             =   1395
      Width           =   4545
      Begin VB.TextBox txtHeightTwips 
         Height          =   285
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "Height in Twips"
         Top             =   2835
         Width           =   750
      End
      Begin VB.TextBox txtWidthTwips 
         Height          =   285
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Width in Twips"
         Top             =   2445
         Width           =   750
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "&Enable"
         Height          =   435
         Index           =   1
         Left            =   2077
         TabIndex        =   13
         ToolTipText     =   "Send enable window message to hWnd"
         Top             =   1470
         Width           =   915
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "&Disable"
         Height          =   435
         Index           =   0
         Left            =   1110
         TabIndex        =   12
         ToolTipText     =   "Send disable window message to hWnd"
         Top             =   1470
         Width           =   915
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C8D0D4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   2820
         ScaleHeight     =   44
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   104
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2445
         Width           =   1590
         Begin VB.Label lblPixel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Current Pixel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   15
            TabIndex        =   27
            Top             =   225
            Width           =   1530
         End
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "Height in Pixels"
         Top             =   2835
         Width           =   750
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Width in Pixels"
         Top             =   2445
         Width           =   750
      End
      Begin VB.TextBox txtBottom 
         Height          =   285
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Bottom"
         Top             =   2055
         Width           =   750
      End
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Right"
         Top             =   2055
         Width           =   750
      End
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Top"
         Top             =   2055
         Width           =   750
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Left"
         Top             =   2055
         Width           =   750
      End
      Begin VB.CommandButton cmdCaption 
         Caption         =   "Change Ca&ption"
         Height          =   435
         Left            =   3045
         TabIndex        =   14
         ToolTipText     =   "Enter new caption in caption textbox and click to change"
         Top             =   1470
         Width           =   1380
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   11
         ToolTipText     =   "Window Caption"
         Top             =   1065
         Width           =   3315
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Class Name"
         Top             =   660
         Width           =   3315
      End
      Begin VB.TextBox txtHwndHex 
         Height          =   285
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "hWnd Hex"
         Top             =   255
         Width           =   1620
      End
      Begin VB.TextBox txtHwndLng 
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "hWnd Long"
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         Height          =   195
         Left            =   555
         TabIndex        =   23
         Top             =   2880
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rect:"
         Height          =   195
         Left            =   675
         TabIndex        =   15
         Top             =   2100
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1110
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hWnd:"
         Height          =   195
         Left            =   585
         TabIndex        =   5
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find Window Information"
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   165
      Width           =   4545
      Begin VB.PictureBox picFinder 
         AutoSize        =   -1  'True
         Height          =   480
         Left            =   1155
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Specify the window by dragging the Finder Tool over a window to select it, then releasing the mouse button when ready."
         Height          =   780
         Left            =   1965
         TabIndex        =   3
         Top             =   255
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finder Tool:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   310
      Y1              =   3
      Y2              =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   310
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Menu mnuFileMain 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Flash Window"
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Enable &Hilite"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Show Zoom"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   4
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptionsMain 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "Zoom 1:1"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Zoom 2:1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Zoom 3:1"
         Index           =   2
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Zoom 4:1"
         Index           =   3
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Zoom 5:1"
         Index           =   4
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Clear Clipboard at Shutdown"
         Checked         =   -1  'True
         Index           =   6
      End
   End
   Begin VB.Menu mnuCaptureMain 
      Caption         =   "&Capture"
      Begin VB.Menu mnuCapture 
         Caption         =   "&Screen Shot"
         Index           =   0
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Selected &Window"
         Index           =   1
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "&Zoom Window"
         Index           =   2
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "&Area"
         Index           =   3
         Begin VB.Menu mnuArea 
            Caption         =   "16 x 16"
            Index           =   0
         End
         Begin VB.Menu mnuArea 
            Caption         =   "32 x 32"
            Index           =   1
         End
         Begin VB.Menu mnuArea 
            Caption         =   "48 x 48"
            Index           =   2
         End
         Begin VB.Menu mnuArea 
            Caption         =   "Drag Custom Area"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "To Clipboard"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "To File"
         Index           =   6
      End
   End
   Begin VB.Menu mnuColorMain 
      Caption         =   "Co&lor"
      Begin VB.Menu mnuColor 
         Caption         =   "Long"
         Index           =   0
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Hex"
         Index           =   1
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Web"
         Index           =   2
      End
      Begin VB.Menu mnuColor 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Red"
         Index           =   4
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Green"
         Index           =   5
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Blue"
         Index           =   6
      End
      Begin VB.Menu mnuColor 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuColor 
         Caption         =   "VB Code"
         Index           =   8
      End
   End
   Begin VB.Menu mnuUtilsMain 
      Caption         =   "&Utils"
      Begin VB.Menu mnuUtils 
         Caption         =   "&Generate GUID"
         Index           =   0
      End
      Begin VB.Menu mnuUtils 
         Caption         =   "&Find Window Code Gen"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'Module:        frmMain
'Description:   UI for Quick Spy, the grunt of the application!
'Date:          May 11, 2001
'Last Updated:  May 16, 2001
'Developer:     Shannon Harmon (shannonh@theharmonfamily.com)
'Info:          Feel free to use any of the routines in your
'               own package, but you may not distribute/sell
'               Quick Spy in whole or part without my consent.
'               Not in code or executable form!
'               I am in no way responsible for any harm you
'               or your friends, your computer, your lunch
'               or cat encounter by using this application:)
'               Copyright 2001, Shannon Harmon - All rights reserved!
'-------------------------------------------------------------
'Notice:        This is a work in progress, I didn't comment most
'               of anything because I wrote this as a utility
'               for me.  I am going to add more stuff to it
'               as I have time or feel the need to.
'               If you have something you want to add to it, please
'               feel free.  If you add something to it,
'               please email me a copy.  Originally written as
'               a sort of tool to make it quicker for me to
'               get some things done.  It wasn't made to replace
'               Microsoft's Spy++ in any way, I just didn't like
'               having to load that everytime I needed a single
'               hWnd of an application.  If you want to find out
'               more about a window, use Spy++!
'               It should work for most Win 32 systems, built
'               and tested with VB 6 SP5 only!
'-------------------------------------------------------------
'Notes:         I do not know of any bugs at this time, but I
'               don't like refreshing the entire desktop when
'               doing the hilite when over a window, it makes
'               a tiny flicker that's annoying.  But when I drew
'               directly on each individual hdc of the window
'               some would not redraw correctly.  Also, since
'               I am not 100% finished with this, I didn't write
'               code to save options selected by the user, I'll
'               do that whenever, if you want it now, then go
'               for it....:)  Clicking on one of the color menu
'               items does nothing more than copy the color to
'               the clipboard.  I am in the middle of writing
'               the portion that draws out palette's, color
'               management, etc...  Zoom window needs to store
'               the entire area that could be used to change
'               from 1:1 to 5:1 so that when user changes it
'               after a zoom is visible, it redraws instead of
'               clearing.  Haven't written that part yet.
'*************************************************************
Option Explicit

Dim blnFind As Boolean
Dim lLastHwnd As Long
Dim dZoomRate As Double
Dim aZoomRates(4) As Double

Private Sub Form_Load()
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  picFinder.BorderStyle = vbBSNone
  picFinder.Picture = LoadResPicture(101, vbResIcon)
  Me.MouseIcon = LoadResPicture(101, vbResCursor)
  lblPixel.Caption = ""
  aZoomRates(0) = 1
  aZoomRates(1) = 2
  aZoomRates(2) = 3
  aZoomRates(3) = 4
  aZoomRates(4) = 5
  dZoomRate = 2
  Call UpdateColorMenuItems
End Sub

Private Sub Form_Resize()
  If mnuFile(2).Checked Then
    If Me.WindowState = vbMinimized Then
      frmZoom.Hide
    Else
      frmZoom.Show
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If mnuOptions(6).Checked And Clipboard.GetFormat(vbCFBitmap) Then Clipboard.Clear
  Set picZoomSave = LoadPicture()
  Dim oForm As Form
  For Each oForm In Forms
    If oForm.Name <> Me.Name Then
      Unload oForm
      Set oForm = Nothing
    End If
  Next
End Sub

Private Sub cmdCaption_Click()
  If lLastHwnd <> 0 Then
    If IsWindow(lLastHwnd) = 0 Then
      MsgBox LoadResString(ERROR_INVALID_WINDOW), vbInformation
    ElseIf GetWindowThreadProcessId(lLastHwnd, 0&) = App.ThreadID Then
      MsgBox LoadResString(ERROR_CANNOT_USE_WITHIN_APP), vbExclamation
    Else
      Call SetWindowText(CLng(txtHwndLng.Text), txtCaption.Text)
    End If
  Else
    MsgBox LoadResString(ERROR_NO_WINDOW_SELECTED), vbInformation
  End If
End Sub

Private Sub cmdEnable_Click(Index As Integer)
  If lLastHwnd <> 0 Then
    If IsWindow(lLastHwnd) = 0 Then
      MsgBox LoadResString(ERROR_INVALID_WINDOW), vbInformation
    ElseIf GetWindowThreadProcessId(lLastHwnd, 0&) = App.ThreadID Then
      MsgBox LoadResString(ERROR_CANNOT_USE_WITHIN_APP), vbExclamation
    Else
      Call EnableWindow(lLastHwnd, Index)
    End If
  Else
    MsgBox LoadResString(ERROR_NO_WINDOW_SELECTED), vbInformation
  End If
End Sub

Private Sub mnuFileMain_Click()
  Dim blnEnabled As Boolean
  If IsWindow(lLastHwnd) = 0 Then
    blnEnabled = False
  Else
    blnEnabled = True
  End If
  
  mnuFile(0).Enabled = blnEnabled
End Sub

Private Sub mnuFile_Click(Index As Integer)
  Select Case Index
    Case 0  'Flash Window
      If lLastHwnd <> 0 Then
        If IsWindow(lLastHwnd) = 0 Then
          MsgBox LoadResString(ERROR_INVALID_WINDOW), vbInformation
        Else
          Dim dc As Long
          Dim hwndDT As Long
          Dim rc As RECT
          hwndDT = GetDesktopWindow()
          dc = GetWindowDC(hwndDT)
          If dc = 0 Then Exit Sub
          Call GetWindowRect(lLastHwnd, rc)
          Call DrawRect(dc, rc, 3, vbBlack): Pause 500
          Call DrawRect(dc, rc, 3, vbRed): Pause 500
          Call DrawRect(dc, rc, 3, vbBlack): Pause 500
          RedrawWindow GetDesktopWindow(), ByVal 0&, ByVal 0&, RDW_FLAGS
          Call ReleaseDC(hwndDT, dc)
        End If
      Else
        MsgBox LoadResString(ERROR_NO_WINDOW_SELECTED), vbInformation
      End If
      
    Case 1  'Enable Hi-Lite
      mnuFile(1).Checked = Not mnuFile(1).Checked
    
    Case 2  'Show Zoom
      mnuFile(2).Checked = Not mnuFile(2).Checked
      If mnuFile(2).Checked Then
        Load frmZoom
        With frmZoom
          .Caption = "Zoom " & dZoomRate & ":1"
          .Width = Screen.TwipsPerPixelX * 250
          .Height = Screen.TwipsPerPixelY * 250
          .Left = Me.Left + IIf(Me.Left > (Screen.Width / 2), -frmZoom.Width, Me.Width)
          .Top = Me.Top
          Set .Picture = picZoomSave
          DoEvents
          .Show
        End With
      Else
        Unload frmZoom
        Set frmZoom = Nothing
      End If
    
    Case 4  'Exit
      Unload Me
  End Select
End Sub

Private Sub mnuOptions_Click(Index As Integer)
  Dim dOldZoomRate As Double
  dOldZoomRate = dZoomRate
  
  Select Case Index
    Case 0, 1, 2, 3, 4  'Zoom rates
      dZoomRate = aZoomRates(Index)
      Dim i As Integer
      For i = 0 To 4
        If aZoomRates(i) = dZoomRate Then
          mnuOptions(i).Checked = True
        Else
          mnuOptions(i).Checked = False
        End If
      Next
  
      If mnuFile(2).Checked Then
        If dZoomRate <> dOldZoomRate Then
          frmZoom.Caption = "Zoom " & dZoomRate & ":1"
          If lLastHwnd <> 0 Then frmZoom.Picture = LoadPicture()
        End If
      End If
    
    Case 6  'Clear bitmap from clipboard on exit
      mnuOptions(6).Checked = Not mnuOptions(6).Checked
  End Select
End Sub

Private Sub mnuCaptureMain_Click()
  mnuCapture(2).Enabled = mnuFile(2).Checked
  If IsWindow(lLastHwnd) = 0 Then
    mnuCapture(1).Enabled = False
  Else
    mnuCapture(1).Enabled = True
  End If
End Sub

Private Sub mnuCapture_Click(Index As Integer)
  Dim strfile As String
  
  Select Case Index
    Case 0, 1
      If Index = 1 And lLastHwnd = 0 Then
        MsgBox LoadResString(ERROR_NO_WINDOW_SELECTED), vbInformation
        Exit Sub
      ElseIf Index = 1 And IsWindow(lLastHwnd) = 0 Then
        MsgBox LoadResString(ERROR_INVALID_WINDOW), vbInformation
        Exit Sub
      End If
      
      Dim dc As Long
      Dim lhWnd As Long
      Dim rc As RECT
      Dim pic As Picture
      
      If Index = 0 Then
        lhWnd = GetDesktopWindow()
      Else
        lhWnd = lLastHwnd
      End If
      
      dc = GetWindowDC(lhWnd)
      Call GetWindowRect(lhWnd, rc)
      
      If Index = 1 Then
        rc.Right = rc.Right - rc.Left
        rc.Bottom = rc.Bottom - rc.Top
        rc.Left = 0
        rc.Top = 0
      Else
        Me.Visible = False
        If mnuFile(2).Checked Then frmZoom.Visible = False
        DoEvents
      End If
      
      If CaptureRect(dc, rc, pic) = 1 Then
        If mnuCapture(5).Checked Then
          Clipboard.Clear
          Clipboard.SetData pic, vbCFBitmap
        Else
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
        MsgBox LoadResString(ERROR_NO_CAPTURE), vbInformation
      End If
      
      If Index = 0 Then
        If mnuFile(2).Checked Then frmZoom.Visible = True
        Me.Visible = True
      End If
      
      Call ReleaseDC(lhWnd, dc)
    
    Case 2
      If mnuFile(2).Checked Then
        If mnuCapture(5).Checked Then
          Clipboard.Clear
          Clipboard.SetData frmZoom.Image, vbCFBitmap
        Else
          strfile = GetSaveFileName(Me.hwnd)
          
          If strfile <> "" Then
            On Error Resume Next
            DoEvents
            SavePicture frmZoom.Picture, strfile
            If Err Then
              MsgBox LoadResString(ERROR_SAVING_IMAGE) & vbCrLf & vbCrLf & _
                     "Number: " & Err.Number & vbCrLf & Err.Description, vbCritical
              Err.Clear
            End If
          End If
        End If
      End If
    
    Case 5
      mnuCapture(5).Checked = True
      mnuCapture(6).Checked = False
    
    Case 6
      mnuCapture(5).Checked = False
      mnuCapture(6).Checked = True
  End Select
End Sub

Private Sub mnuArea_Click(Index As Integer)
  Me.WindowState = vbMinimized
  DoEvents
  
  Select Case Index
    Case 0: CA_Cur = CA_16
    Case 1: CA_Cur = CA_32
    Case 2: CA_Cur = CA_48
    Case 3: CA_Cur = CA_CUSTOM
  End Select
  
  frmCaptureArea.Show vbModal, Me
  Me.WindowState = vbNormal
  Set frmCaptureArea = Nothing
End Sub

Private Sub mnuColor_Click(Index As Integer)
  Dim lCol As Long
  lCol = picColor.BackColor
  
  Clipboard.Clear
  
  Select Case Index
    Case 0  'Long
      Clipboard.SetText CStr(lCol)
    Case 1  'Hex
      Clipboard.SetText "&H" & CStr(Hex(lCol))
    Case 2  'Web
      Clipboard.SetText WebColor(lCol)
    Case 4  'Red
      Clipboard.SetText CStr(RGBRed(lCol))
    Case 5  'Green
      Clipboard.SetText CStr(RGBGreen(lCol))
    Case 6  'Blue
      Clipboard.SetText CStr(RGBBlue(lCol))
    Case 8  'VB Color
      Clipboard.SetText GetVBColorName(lCol)
  End Select
End Sub

Private Sub mnuUtilsMain_Click()
  Dim blnEnabled As Boolean
  If IsWindow(lLastHwnd) = 0 Then
    blnEnabled = False
  Else
    blnEnabled = True
  End If
  
  mnuUtils(1).Enabled = blnEnabled
End Sub

Private Sub mnuUtils_Click(Index As Integer)
  Select Case Index
    Case 0  'Generate GUID
      frmGUID.Show vbModal, Me
    
    Case 1  'Find Window Code Gen
      If mnuFile(2).Checked Then frmZoom.Visible = False
      Me.Visible = False
      frmFWCode.Show
  End Select
End Sub

Private Sub picColor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuColorMain
  End If
End Sub

Private Sub lblPixel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuColorMain
  End If
End Sub

Private Sub picFinder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  blnFind = True
  picFinder.Picture = LoadResPicture(102, vbResIcon)
  Me.MousePointer = 99
  If mnuFile(2).Checked Then frmZoom.AutoRedraw = False
End Sub

Private Sub picFinder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If blnFind Then
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim dc As Long
    Dim lhWnd As Long
    Dim strBuf As String
    
    Call GetCursorPos(pt)
    lhWnd = WindowFromPoint(pt.x, pt.y)
    
    'Don't process till out of finder picturebox
    If lhWnd = picFinder.hwnd Then Exit Sub
    
    'Uncomment this section to exclude any window within this app's thread from processing
    'If you do, you will have to write code to compliment it in the mouseup event.
    'If GetWindowThreadProcessId(lhWnd, 0&) = App.ThreadID Then Exit Sub
    
    dc = GetWindowDC(GetDesktopWindow())
    Call GetWindowRect(lhWnd, rc)
    
    If lLastHwnd <> lhWnd And lLastHwnd <> 0 And mnuFile(1).Checked Then
      Call RedrawWindow(GetDesktopWindow(), ByVal 0&, ByVal 0&, RDW_FLAGS)
    End If
    
    If mnuFile(1).Checked Then Call DrawRect(dc, rc, 2)
    
    txtHwndLng.Text = lhWnd
    txtHwndHex.Text = "&H" & CStr(Hex(lhWnd))
    
    strBuf = Space(256)
    Call GetClassName(lhWnd, strBuf, 256)
    txtClass = strBuf
    
    strBuf = String$(255, vbNullChar)
    Call GetWindowText(lhWnd, strBuf, 255)
    txtCaption.Text = strBuf
    
    With rc
      txtLeft.Text = .Left
      txtTop.Text = .Top
      txtRight.Text = .Right
      txtBottom.Text = .Bottom
      txtWidth.Text = .Right - .Left
      txtWidthTwips = Screen.TwipsPerPixelX * (.Right - .Left)
      txtHeight.Text = .Bottom - .Top
      txtHeightTwips = Screen.TwipsPerPixelY * (.Bottom - .Top)
    End With
    
    picColor.BackColor = GetPixel(dc, pt.x, pt.y)
    lblPixel.ForeColor = IIf(picColor.BackColor < 8388607, vbWhite, vbBlack)
    lblPixel.Caption = "X: " & pt.x & ", Y: " & pt.y
    
    If mnuFile(2).Checked Then
      If dZoomRate = 1 Then
        Call BitBlt(frmZoom.hdc, 0, 0, frmZoom.ScaleWidth, frmZoom.ScaleHeight, dc, _
        pt.x - CLng(frmZoom.ScaleWidth / 2), pt.y - CLng(frmZoom.ScaleHeight / 2), SRCCOPY)
      Else
        Dim iPix As Integer
        iPix = CInt(frmZoom.ScaleHeight / dZoomRate)
        Call StretchBlt(frmZoom.hdc, 0, 0, frmZoom.ScaleWidth, frmZoom.ScaleHeight, dc, _
        pt.x - CLng(iPix / 2), pt.y - CLng(iPix / 2), iPix, iPix, SRCCOPY)
      End If
      
      Call BitBlt(frmZoom.picSave.hdc, 0, 0, frmZoom.ScaleWidth, _
      frmZoom.ScaleHeight, frmZoom.hdc, 0, 0, SRCCOPY)
    End If
    
    Call ReleaseDC(GetDesktopWindow(), dc)
    lLastHwnd = lhWnd
  End If
End Sub

Private Sub picFinder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  blnFind = False
  picFinder.Picture = LoadResPicture(101, vbResIcon)
  Me.MousePointer = 0
  Call UpdateColorMenuItems
    
  If mnuFile(1).Checked Then
    Call RedrawWindow(GetDesktopWindow(), ByVal 0&, ByVal 0&, RDW_FLAGS)
  End If
  
  If mnuFile(2).Checked Then
    frmZoom.AutoRedraw = True
    frmZoom.Picture = frmZoom.picSave.Image
    DoEvents
    Set picZoomSave = frmZoom.Picture
  End If
End Sub

Private Sub UpdateColorMenuItems()
  Dim lCol As Long
  lCol = picColor.BackColor
  mnuColor(0).Caption = "Long" & vbTab & lCol
  mnuColor(1).Caption = "Hex" & vbTab & "&&H" & Hex(lCol)
  mnuColor(2).Caption = "Web" & vbTab & "#" & WebColor(lCol)
  mnuColor(4).Caption = "Red" & vbTab & RGBRed(lCol)
  mnuColor(5).Caption = "Green" & vbTab & RGBGreen(lCol)
  mnuColor(6).Caption = "Blue" & vbTab & RGBBlue(lCol)
  mnuColor(8).Caption = "VB Code" & vbTab & GetVBColorName(lCol, "N/A")
End Sub
