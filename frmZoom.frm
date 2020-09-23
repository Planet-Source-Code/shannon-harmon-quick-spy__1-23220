VERSION 5.00
Begin VB.Form frmZoom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Zoom"
   ClientHeight    =   2250
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   90
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuPopupMain 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "&Copy to Clipboard"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'Module:        frmZoom
'Description:   Zoom window for Quick Spy
'Date:          May 11, 2001
'Last Updated:  May 16, 2001
'Developer:     Shannon Harmon (shannonh@theharmonfamily.com)
'Info:          Feel free to use any of the routines in your
'               own package, but you may not distribute/sell
'               Quick Spy in whole or part without my consent.
'               Copyright 2001, Shannon Harmon - All rights reserved!
'*************************************************************
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyS And Shift = vbCtrlMask Then Unload Me
End Sub

Private Sub Form_Load()
  Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    picSave.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuFile(2).Checked = False
End Sub

