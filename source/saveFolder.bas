Attribute VB_Name = "saveFolder"
Option Explicit

Public Const MAX_PATH = 260

'common to both methods
Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
   (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public Declare Function SHGetSpecialFolderLocation _
   Lib "shell32.dll" _
  (ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   pidl As Long) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
    
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1

Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)


Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Public Declare Function SHSimpleIDListFromPath Lib _
   "shell32" Alias "#162" _
   (ByVal szPath As String) As Long


'specific to the STRING method
Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" _
   (lpString1 As Any, lpString2 As Any) As Long

Public Declare Function lstrlenA Lib "kernel32" _
   (lpString As Any) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'windows-defined type OSVERSIONINFO
Public Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type
        
Public Const VER_PLATFORM_WIN32_NT = 2

  
Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
   Alias "GetLogicalDriveStringsA" _
  (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long



Public Function BrowseCallbackProcStr(ByVal hwnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hwnd, BFFM_SETSELECTIONA, _
                          1&, ByVal lpData)
                          
         Case Else:
         
   End Select
          
End Function
          

Public Function BrowseCallbackProc(ByVal hwnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hwnd, BFFM_SETSELECTIONA, _
                          0&, ByVal lpData)
                          
         Case Else:
         
   End Select

End Function


Public Function FARPROC(pfn As Long) As Long

  FARPROC = pfn

End Function
