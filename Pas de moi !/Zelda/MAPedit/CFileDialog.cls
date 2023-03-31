VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Declare Function GetOpenFileName _
    Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" _
    (pOpenfilename As OPENFILENAME) _
    As Long

Private Declare Function GetSaveFileName _
    Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" _
    (pOpenfilename As OPENFILENAME) _
    As Long

Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strInitialDir As String
Private m_strFilter As String
Private m_intFilterIndex As Integer
Private m_intMaxFileSize As Integer
Private m_lnghWndParent As Long

Private Const cintMaxFileLength As Integer = 260

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
    of.Flags = &H80000
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
    
    strTemp = m_strFileName & String$(cintMaxFileLength - Len(m_strFileName), 0)
    of.lpstrFile = strTemp
    of.nMaxFile = cintMaxFileLength
    
    strTemp = m_strFileTitle & String$(cintMaxFileLength - Len(m_strFileTitle), 0)
    of.lpstrFileTitle = strTemp
    of.lpstrInitialDir = m_strInitialDir
    of.nMaxFileTitle = cintMaxFileLength
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
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Show"
    Resume PROC_EXIT

End Function

Private Function TrimNulls(ByVal strIn As String) As String
    Dim intPos As Integer
    
    On Error GoTo PROC_ERR
    
    intPos = InStr(strIn, vbNullChar)
    
    If intPos = 0 Then
    TrimNulls = strIn
    Else
    If intPos = 1 Then
        TrimNulls = ""
    Else
        TrimNulls = Left$(strIn, intPos - 1)
    End If
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TrimNulls"
    Resume PROC_EXIT
    
End Function


