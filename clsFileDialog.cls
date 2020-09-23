VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************
'Module:        clsFileDialog
'Description:   Replacement for the Common Dialog control's
'               file save/open library
'*************************************************************
Option Explicit

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strInitialDir As String
Private m_strFilter As String
Private m_intFilterIndex As Integer
Private m_eFlags As EnumFilFlags
Private m_intMaxFileSize As Integer
Private m_lnghWndParent As Long

Private Const cintMaxFileLength As Integer = 260

Public Enum EnumFilFlags
  FleReadOnly = &H1
  FleOverWritePrompt = &H2
  FleHideReadOnly = &H4
  FleNoChangeDir = &H8
  FleShowHelp = &H10
  FleEnableHook = &H20
  FleEnableTemplate = &H40
  FleEnableTemplateHandle = &H80
  FleNoValidate = &H100
  FleAllowMultiSelect = &H200
  FleExtensionDifferent = &H400
  FlePathMustExist = &H800
  FleFileMustExist = &H1000
  FleCreatePrompt = &H2000
  FleShareAware = &H4000
  FleNoReadOnlyReturn = &H8000
  FleNoTestFileCreate = &H10000
  FleNoNetworkButton = &H20000
  FleExplorer = &H80000
  FleLongnames = &H200000
End Enum

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Class_Initialize()
  m_intMaxFileSize = cintMaxFileLength
End Sub

Public Property Get DefaultExt() As String
  DefaultExt = m_strDefaultExt
End Property

Public Property Let DefaultExt(ByVal strValue As String)
  m_strDefaultExt = strValue
End Property

Public Property Get DialogTitle() As String
  DialogTitle = m_strDialogTitle
End Property

Public Property Let DialogTitle(ByVal strValue As String)
  m_strDialogTitle = strValue
End Property

Public Property Get FileName() As String
  FileName = m_strFileName
End Property

Public Property Let FileName(ByVal strValue As String)
  m_strFileName = strValue
End Property

Public Property Get FileTitle() As String
  FileTitle = m_strFileTitle
End Property

Public Property Let FileTitle(ByVal strValue As String)
  m_strFileTitle = strValue
End Property

Public Property Get Filter() As String
  Filter = m_strFilter
End Property

Public Property Let Filter(ByVal strValue As String)
  m_strFilter = strValue
End Property

Public Property Get FilterIndex() As Integer
  FilterIndex = m_intFilterIndex
End Property

Public Property Let FilterIndex(ByVal intValue As Integer)
  m_intFilterIndex = intValue
End Property

Public Property Get Flags() As EnumFilFlags
  Flags = m_eFlags
End Property

Public Property Let Flags(ByVal eValue As EnumFilFlags)
  m_eFlags = eValue
End Property

Public Property Get hWndParent() As Long
  hWndParent = m_lnghWndParent
End Property

Public Property Let hWndParent(ByVal lngValue As Long)
  m_lnghWndParent = lngValue
End Property

Public Property Get InitialDir() As String
  InitialDir = m_strInitialDir
End Property

Public Property Let InitialDir(ByVal strValue As String)
  m_strInitialDir = strValue
End Property

Public Property Get MaxFileSize() As Integer
  MaxFileSize = m_intMaxFileSize
End Property

Public Property Let MaxFileSize(ByVal intValue As Integer)
  m_intMaxFileSize = intValue
End Property

Public Function Show(fOpen As Boolean) As Boolean
  Dim of As OPENFILENAME
  Dim strChar As String * 1
  Dim intCounter As Integer
  Dim strTemp As String
  On Error GoTo PROC_ERR
  
  of.lpstrTitle = m_strDialogTitle & ""
  of.Flags = m_eFlags
  of.lpstrDefExt = m_strDefaultExt & ""
  of.lStructSize = LenB(of)
  of.lpstrFilter = m_strFilter & "||"
  of.nFilterIndex = m_intFilterIndex
  
  For intCounter = 1 To Len(m_strFilter)
    strChar = Mid$(m_strFilter, intCounter, 1)
    If strChar = "|" Then
      strTemp = strTemp & vbNullChar
    Else
      strTemp = strTemp & strChar
    End If
  Next
  
  strTemp = strTemp & vbNullChar & vbNullChar
  of.lpstrFilter = strTemp
  
  strTemp = m_strFileName & String$(m_intMaxFileSize - Len(m_strFileName), 0)
  of.lpstrFile = strTemp
  of.nMaxFile = m_intMaxFileSize
  
  strTemp = m_strFileTitle & String$(m_intMaxFileSize - Len(m_strFileTitle), 0)
  of.lpstrFileTitle = strTemp
  of.lpstrInitialDir = m_strInitialDir
  of.nMaxFileTitle = m_intMaxFileSize
  of.hwndOwner = m_lnghWndParent
  
  If fOpen Then
    If GetOpenFileName(of) Then
      Show = True
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      Show = False
    End If
  Else
    If GetSaveFileName(of) Then
      Show = True
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      Show = False
    End If
  End If
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "Show"
  Resume PROC_EXIT
End Function

Private Function TrimNulls(ByVal strIn As String) As String
  Dim intPos As Integer
  Dim strTemp As String
  Dim strTrimNullsRet As String
  Dim strChar As String
  Dim intCounter As Integer
  On Error GoTo PROC_ERR
  
  intPos = InStr(strIn, vbNullChar)
  
  Do While intPos > 0
    If (intPos + 1 < Len(strIn)) Then
      If (Mid$(strIn, intPos + 1, 1) = vbNullChar) Then
        Exit Do
      End If
    End If
    intPos = InStr(intPos + 1, strIn, vbNullChar)
  Loop
  
  If intPos = 0 Then
    TrimNulls = strIn
  Else
    If intPos = 1 Then
      strTrimNullsRet = ""
    Else
      strTrimNullsRet = Left$(strIn, intPos - 1)
    End If
  End If
  
  For intCounter = 1 To Len(strTrimNullsRet)
    strChar = Mid$(strTrimNullsRet, intCounter, 1)
    If strChar = vbNullChar Then
      strTemp = strTemp & " "
    Else
      strTemp = strTemp & strChar
    End If
  Next
  
  TrimNulls = strTemp
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "TrimNulls"
  Resume PROC_EXIT
End Function
