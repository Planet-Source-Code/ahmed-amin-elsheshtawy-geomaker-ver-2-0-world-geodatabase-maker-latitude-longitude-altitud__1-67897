Attribute VB_Name = "mBrowseForFolder"
'v2.08b
'http://www.vbxsystem.de/en/?vbx/modules/browseforfolder_vb.htm
'this module is not VBScript compatible

'Add the 'Microsoft Shell Controls and Automation'-Reference

'known issues: with BIF_NEWDIALOGSTYLE, BIF_RETURNONLYFSDIRS doesn't gray up the OK button _
'if the user selects a non-file system object, the return value is an empty string (same as canceled)

Option Explicit

'BrowseInfo struct
Private Type BROWSEINFO
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

'buffer for return
Private Const cintMaxSize = 512

'Ausgangsordner
Private mstrDefaultPath As String

'API functions
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hWndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)

'main function
Public Function BrowseForFolder(Optional strTitle As String = "", Optional strDefaultPath As String = "", Optional shRootFolder As ShellSpecialFolderConstants = ssfDRIVES, Optional bolEditBox As Boolean = False, Optional bolIncludeFiles As Boolean = False, Optional frmParent As Form = Nothing) As String
   
   Dim tBrowseInfo As BROWSEINFO
   Dim lnghWndParent As Long
   Dim lngOptions  As Long
   Dim lngRootPIDL As Long
   Dim lngPIDL As Long
   Dim strPath As String
   
   'Parent form
   If Not frmParent Is Nothing Then
      lnghWndParent = frmParent.hWnd
   End If
   
   mstrDefaultPath = strDefaultPath
   
   'BIF_RETURNONLYFSDIRS, BIF_NEWDIALOGSTYLE
   lngOptions = &H1 Or &H40
   'BIF_BROWSEINCLUDEFILES
   If bolIncludeFiles Then lngOptions = lngOptions Or &H4000
   'BIF_EDITBOX
   If bolEditBox Then lngOptions = lngOptions Or &H10
   
   'PIDL of the parent folder
   SHGetSpecialFolderLocation lnghWndParent, shRootFolder, lngRootPIDL
   
   'Fill struct
   With tBrowseInfo
      .hWndOwner = lnghWndParent
      .pszDisplayName = Space$(cintMaxSize)
      .pIDLRoot = lngRootPIDL
      .ulFlags = lngOptions
      .lpszTitle = strTitle
      .lpfn = getProcAddress(AddressOf BrowseForFolderCallBack)
   End With
      
   'Execute BrowseForFolder
   lngPIDL = SHBrowseForFolder(tBrowseInfo)
   
   'Get folder by PIDL
   If lngPIDL = 0 Then
      strPath = ""
   Else
      strPath = Space$(cintMaxSize)
      SHGetPathFromIDList lngPIDL, strPath
      strPath = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
   End If
   BrowseForFolder = strPath
   
   'Free ressources
   CoTaskMemFree lngPIDL
End Function

'returns the location (path) of a special foldr by the ShellSpecialFolderConstant
Public Function GetSpecialFolderLocation(shFolder As ShellSpecialFolderConstants) As String
    Dim lngPIDL
    Dim strPath As String
    strPath = Space$(cintMaxSize)
    SHGetSpecialFolderLocation 0, shFolder, lngPIDL
    SHGetPathFromIDList lngPIDL, strPath
    GetSpecialFolderLocation = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
End Function

'Sets the predefined folder
Private Function BrowseForFolderCallBack(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   '1- Init
   If uMsg = 1 Then
      'BFFM_SETSELECTION
      If mstrDefaultPath <> vbNullString Then SendMessage hWnd, &H400 + 102, True, ByVal mstrDefaultPath
   End If
End Function

'Address of the procedure
Private Function getProcAddress(ByVal lngProcAddress As Long)
   getProcAddress = lngProcAddress
End Function

'2.08c - added getSpecialFolderLocation function
'2.08b - made some changes to fit the 2.08 conventions
'          - removed SpecialFolderConstants as they are declared in 'Microsoft Shell Controls and Automation'

'Philipp Fabrizio - ©2002 VBXSystem.de

