Attribute VB_Name = "FileRename_Module_Main"
'Module1:
Option Explicit

Public r_output As New ADODB.Recordset
Public r_FileInfo As New ADODB.Recordset

Public g_CurDirectory As String    '当前选中的目录








Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Const LPTR = (&H0 Or &H40)
Public Type BrowseInfo
    hWndOwner             As Long
    pIDLRoot             As Long
    pszDisplayName   As Long
    lpszTitle             As Long
    ulFlags                 As Long
    lpfnCallback     As Long
    lParam                 As Long
    iImage                 As Long
End Type







Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const BFFM_SETSelectIONA       As Long = (WM_USER + 102)
Private Const BFFM_SETSelectIONW       As Long = (WM_USER + 103)
Private Const BFFM_INITIALIZED       As Long = 1



' ============================================================
' 模块名称：
'
' 模块描述：
'
'     获取文件的日期信息
'
'
'
'
' ============================================================

'Declaratins:
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800

Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type

Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long





Public Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)  ' Get information of file using API call
    If plngFirstFileHwnd = 0 Then
      Findfile.cFileName = "Error"                              ' If file was not found, return error as name
    Else
      Findfile = Win32Data                                      ' Else return results
    End If
    plngRtn = FindClose(plngFirstFileHwnd)                      ' It is important that you close the handle for FindFirstFile
End Function


   Public Function BrowseForFolders_CallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
        If uMsg = BFFM_INITIALIZED Then
            SendMessage hWnd, BFFM_SETSelectIONA, True, ByVal lpData
        End If
   End Function


