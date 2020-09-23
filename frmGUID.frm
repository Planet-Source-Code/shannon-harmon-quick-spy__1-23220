VERSION 5.00
Begin VB.Form frmGUID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate GUID"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmGUID.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   420
      Left            =   810
      TabIndex        =   3
      ToolTipText     =   "Copy to clipboard"
      Top             =   3840
      Width           =   1110
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   420
      Left            =   2025
      TabIndex        =   4
      ToolTipText     =   "Click to Gernerate GUID(s)"
      Top             =   3840
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   420
      Left            =   3240
      TabIndex        =   5
      Top             =   3840
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Global Unique Identifier"
      Height          =   3570
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   4290
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "1"
         Top             =   3135
         Width           =   810
      End
      Begin VB.TextBox txtGUID 
         Height          =   2685
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   285
         Width           =   4065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty to Generate:"
         Height          =   195
         Left            =   2130
         TabIndex        =   1
         Top             =   3180
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmGUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'Module:        frmGUID
'Description:   GUID Generator Utility for Quick Spy
'Date:          May 11, 2001
'Last Updated:  May 16, 2001
'Developer:     Shannon Harmon (shannonh@theharmonfamily.com)
'Info:          Feel free to use any of the routines in your
'               own package, but you may not distribute/sell
'               Quick Spy in whole or part without my consent.
'               Copyright 2001, Shannon Harmon - All rights reserved!
'*************************************************************
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtGUID.Text
End Sub

Private Sub cmdGenerate_Click()
  If IsNumeric(txtQty) Then
    Dim i As Integer
    Dim x As Integer
    Dim strGUIDs As String
    
    x = CInt(txtQty)
    strGUIDs = ""
    
    Screen.MousePointer = vbHourglass
    
    For i = 1 To x
      strGUIDs = strGUIDs & GetGUID()
      If i <> x Then strGUIDs = strGUIDs & vbNewLine
    Next i
      
    txtGUID.Text = strGUIDs
    
    Screen.MousePointer = vbNormal
  End If
End Sub

Private Sub txtGUID_GotFocus()
  txtGUID.SelStart = 0
  txtGUID.SelLength = Len(txtGUID.Text)
End Sub

Private Sub txtQty_Change()
  If Trim(txtQty) = "" Then
    If cmdGenerate.Enabled Then cmdGenerate.Enabled = False
  Else
    If Not cmdGenerate.Enabled Then cmdGenerate.Enabled = True
  End If
End Sub

Private Sub txtQty_GotFocus()
  txtQty.SelStart = 0
  txtQty.SelLength = Len(txtQty)
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Or KeyCode = vbKeyAdd Then
    If txtQty.Text = "" Then
      txtQty.Text = "1"
    Else
      If txtQty.Text <> "999" Then txtQty.Text = CInt(txtQty.Text) + 1
    End If
  ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeySubtract Then
    If txtQty.Text = "" Then
      txtQty.Text = "1"
    Else
      If txtQty.Text <> "1" Then txtQty.Text = CInt(txtQty.Text) - 1
    End If
  ElseIf KeyCode = vbKeyPageUp Then
    txtQty.Text = "999"
  ElseIf KeyCode = vbKeyPageDown Then
    txtQty.Text = "1"
  End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then
    Exit Sub
  ElseIf Not IsNumeric(Chr$(KeyAscii)) Then
    KeyAscii = 0
  End If
End Sub
