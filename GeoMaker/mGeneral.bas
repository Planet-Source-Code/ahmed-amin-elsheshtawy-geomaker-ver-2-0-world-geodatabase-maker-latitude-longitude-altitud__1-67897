Attribute VB_Name = "mGeneral"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft GeoMaker
'Program Author   : Dr. Elsheshtawy, Ahmed Amin, Ph.D.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2007-2009 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit

Public Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
   dwTotalPhys As Long
   dwAvailPhys As Long
   dwTotalPageFile As Long
   dwAvailPageFile As Long
   dwTotalVirtual As Long
   dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'The GetTickCount function retrieves the number of milliseconds that have elapsed since the system was started. It is limited to the resolution of the system timer.
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function CopyFileAPI Lib "kernel32" _
                Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                ByVal lpNewFileName As String, ByVal bFailIfExists As Long) _
                As Long

Private Const MaxPathName = 256
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(MaxPathName) As Byte
End Type

Private Const OF_SHARE_EXCLUSIVE = &H10
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const MAX_PATH = 260
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const ERROR_NO_MORE_FILES = 18&

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" Alias _
   "FindNextFileA" (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" _
   (ByVal hFindFile As Long) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

'====================================================================
'====================================================================

' Sub to sleep x milli seconds
Public Sub Sleep(lngSleep As Long)
   Dim lngSleepEnd As Long
   'GetTickCount: Retrieves the number of milliseconds that have elapsed since the system was started, up to 49.7 days.
   lngSleepEnd = GetTickCount + lngSleep '* 1000
   While GetTickCount <= lngSleepEnd
      DoEvents
   Wend
End Sub

Public Function TrimNull(sString As String) As String
    TrimNull = Left(sString, InStr(1, sString, vbNullChar) - 1)
End Function

'Sort string arrays
Public Sub SortArray(inpArray() As String)
    
    Dim intRet
    Dim intCompare
    Dim intLoopTimes
    Dim strTemp
    
    For intLoopTimes = 1 To UBound(inpArray)
        For intCompare = LBound(inpArray) To UBound(inpArray) - 1
            intRet = StrComp(inpArray(intCompare), _
                     inpArray(intCompare + 1), vbTextCompare)
    
            If intRet = 1 Then 'String1 is greater than String2
                strTemp = inpArray(intCompare)
                inpArray(intCompare) = inpArray(intCompare + 1)
                inpArray(intCompare + 1) = strTemp
            End If
        Next
    Next

End Sub

' For put a windows in the middle of the screen
' FrmChild  = Windows to center
' FrmParent = MDI Windows (Optional)
Public Sub CenterForm(FrmChild As Form, Optional frmParent As Variant)
    Dim iTop As Integer, iLeft As Integer
    
    If Not IsMissing(frmParent) Then
        iTop = frmParent.Top + (frmParent.ScaleHeight - FrmChild.Height) \ 2
        iLeft = frmParent.Left + (frmParent.ScaleWidth - FrmChild.Width) \ 2
    Else
        iTop = (Screen.Height - FrmChild.Height) \ 2
        iLeft = (Screen.Width - FrmChild.Width) \ 2
    End If
    If iTop And iLeft Then
        FrmChild.Move iLeft, iTop
    End If
End Sub

Function NoNulo(Vrx As Variant) As String
    If IsNull(Vrx) Then
        NoNulo = ""
    Else
        NoNulo = Vrx
    End If
End Function

Public Function DirExists(ByVal sDir As String) As Boolean

    On Error GoTo Err_Handler
    Dim strDir As String

    strDir = Dir(sDir, vbDirectory)

    If (strDir = "") Then
         'If it doesn't exist, create it
        CreateDirectoryStruct sDir
    End If
    
    DirExists = True
    Exit Function

Err_Handler:
    DirExists = False
End Function

Public Sub CreateDirectoryStruct(ByVal CreateThisPath As String)

    On Error GoTo Err_Handler
    'do initial check
    Dim Ret As Boolean
    Dim temp As String
    Dim ComputerName As String
    Dim IntoItCount As Integer
    Dim X As Integer
    Dim WakeString As String
    Dim MadeIt As Integer

    If Dir$(CreateThisPath, vbDirectory) <> "" Then Exit Sub
    'is this a network path?

    If Left$(CreateThisPath, 2) = "\\" Then ' this is a UNC NetworkPath
        'must extract the machine name first, th
        '     en get to the first folder
        IntoItCount = 3
        ComputerName = Mid$(CreateThisPath, IntoItCount, InStr(IntoItCount, CreateThisPath, "\") - IntoItCount)
        IntoItCount = IntoItCount + Len(ComputerName) + 1
        IntoItCount = InStr(IntoItCount, CreateThisPath, "\") + 1
        'temp = Mid$(CreateThisPath, IntoItCount
        '     , x)
    Else ' this is a regular path
        IntoItCount = 4
    End If
    WakeString = Left$(CreateThisPath, IntoItCount - 1)
    'start a loop through the CreateThisPath
    '     string

    Do
        X = InStr(IntoItCount, CreateThisPath, "\")

        If X <> 0 Then
            X = X - IntoItCount
            temp = Mid$(CreateThisPath, IntoItCount, X)
        Else
            temp = Mid$(CreateThisPath, IntoItCount)
        End If
        IntoItCount = IntoItCount + Len(temp) + 1
        temp = WakeString + temp
        'Create a directory if it doesn't alread
        '     y exist
        Ret = (Dir$(temp, vbDirectory) <> "")


        If Not Ret Then
            'ret& = CreateDirectory(temp, Security)
            MkDir temp
        End If
        IntoItCount = IntoItCount 'track where we are in the String
        WakeString = Left$(CreateThisPath, IntoItCount - 1)
    Loop While WakeString <> CreateThisPath

    Exit Sub

Err_Handler:
    Err.Raise Err.Number
End Sub

Private Sub ClearDirectory(ByVal psDirName As String)
    'This function attempts to delete all fi
    '     les
    'and subdirectories of the given
    'directory name, and leaves the given
    'directory intact, but completely empty.
    '
    '
    'If the Kill command generates an error
    '     (i.e.
    'file is in use by another process -
    'permission denied error), then that fil
    '     e and
    'subdirectory will be skipped, and the
    'program will continue (On Error Resume
'     Next).
'
'EXAMPLE CALL:
' ClearDirectory "C:\Temp\"
Dim sSubDir


If Len(psDirName) > 0 Then


    If Right(psDirName, 1) <> "\" Then
        psDirName = psDirName & "\"
    End If
    'Attempt to remove any files in director
    '     y
    'with one command (if error, we'll
    'attempt to delete the files one at a
    'time later in the loop):
    On Error Resume Next
    Kill psDirName & "*.*"


    DoEvents
        
        sSubDir = Dir(psDirName, vbDirectory)


        Do While Len(sSubDir) > 0
            'Ignore the current directory and the
            'encompassing directory:
            If sSubDir <> "." And _
            sSubDir <> ".." Then
            'Use bitwise comparison to make
            'sure MyName is a directory:
            If (GetAttr(psDirName & sSubDir) And _
            vbDirectory) = vbDirectory Then
            'Use recursion to clear files
            'from subdir:
            ClearDirectory psDirName & _
            sSubDir & "\"
            'Remove directory once files
            'have been cleared (deleted)
            'from it:
            RmDir psDirName & sSubDir


            DoEvents
                'ReInitialize Dir Command
                'after using recursion:
                sSubDir = Dir(psDirName, vbDirectory)
            Else
                'This file is remaining because
                'most likely, the Kill statement
                'before this loop errored out
                'when attempting to delete all
                'the files at once in this
                'directory. This attempt to
                'delete a single file by itself
                'may work because another
                '(locked) file within this same
                'directory may have prevented
                '(non-locked) files from being
                'deleted:
                Kill psDirName & sSubDir
                sSubDir = Dir
            End If
        Else
            sSubDir = Dir
        End If
    Loop
End If
End Sub

'Public Sub ColorListviewRow(lv As ListView, RowNbr As Long, _
'                                RowColor As OLE_COLOR)
''***************************************************************************
''Purpose: Color a ListView Row
''Inputs : lv - The ListView
''         RowNbr - The index of the row to be colored
''         RowColor - The color to color it
''Outputs: None
''***************************************************************************
'    Dim itmX As ListItem
'    Dim lvSI As ListSubItem
'    Dim intIndex As Integer
'
'    On Error GoTo ErrorRoutine
'
'    Set itmX = lv.ListItems(RowNbr)
'
'    itmX.ForeColor = RowColor
'    For intIndex = 1 To lv.ColumnHeaders.Count - 1
'        Set lvSI = itmX.ListSubItems(intIndex)
'        lvSI.ForeColor = RowColor
'    Next
'
'    Set itmX = Nothing
'    Set lvSI = Nothing
'    Exit Sub
'
'ErrorRoutine:
'    'MsgBox "List view set color error: " & Err.Description
'End Sub
'
'Public Sub ColorListviewSubItem(lv As ListView, RowNbr As Long, _
'                                SubItemNbr As Long, RowColor As OLE_COLOR)
''***************************************************************************
''Purpose: Color a ListView Row
''Inputs : lv - The ListView
''         RowNbr - The index of the row to be colored
''         RowColor - The color to color it
''Outputs: None
''***************************************************************************
'    Dim itmX As ListItem
'    Dim lvSI As ListSubItem
'    Dim intIndex As Integer
'
'    On Error GoTo ErrorRoutine
'
'    Set itmX = lv.ListItems(RowNbr)
'    'itmX.ForeColor = RowColor
'
'    Set lvSI = itmX.ListSubItems(SubItemNbr)
'    lvSI.ForeColor = RowColor
'
'    Set itmX = Nothing
'    Set lvSI = Nothing
'    Exit Sub
'
'ErrorRoutine:
'    'MsgBox "List view set color error: " & Err.Description
'End Sub
'
Public Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
        QualifyPath = sPath & "\"
   Else
        QualifyPath = sPath
   End If
      
End Function

Public Function AppPath() As String
    
    Dim sPath As String
    
    sPath = App.Path
   If Right$(sPath, 1) <> "\" Then
        AppPath = sPath & "\"
   Else
        AppPath = sPath
   End If
End Function

Public Function safeBound(Item) As Integer
    On Error GoTo NullBound
    safeBound = UBound(Item)
    Exit Function
NullBound:
    safeBound = -1
End Function

' Return a formatted string representing
' the number of bytes.
Public Function FormatBytes(ByVal Value As Double)
    
    Dim bSize(8) As String
    Dim i As Integer
    Dim b As Double
    
    bSize(0) = "Bytes"
    bSize(1) = "KB" 'Kilobytes
    bSize(2) = "MB" 'Megabytes
    bSize(3) = "GB" 'Gigabytes
    bSize(4) = "TB" 'Terabytes
    bSize(5) = "PB" 'Petabytes
    bSize(6) = "EB" 'Exabytes
    bSize(7) = "ZB" 'Zettabytes
    bSize(8) = "YB" 'Yottabytes
    
    b = CDbl(Value) ' Make sure var is a Double (not just variant)
    For i = UBound(bSize) To 0 Step -1
       If b >= (1024 ^ i) Then
          If i = 0 Then
            FormatBytes = Format((Round(b / (1024 ^ i), 3)), "###,###,###,###") & " " & bSize(i)
          ElseIf i = 1 Then
            FormatBytes = Format((Round(b / (1024 ^ i), 3)), "###,###,###,###0.00") & " " & bSize(i)
          Else
            FormatBytes = Format((Round(b / (1024 ^ i), 3)), "###,###,###,###0.000") & " " & bSize(i)
          End If
          Exit For
       End If
    Next
    
End Function

'Find the default executable for a file
Private Function FindFileExecutable(ByVal FileName As String) As String
    
    Dim FilePath As String
    Dim pos As Integer
    Dim Result As String

    pos = InStrRev(FileName, "\")
    FilePath = Left$(FileName, pos)
    FileName = Mid$(FileName, pos + 1)
    Result = Space$(1024)

    FindExecutable FileName, FilePath, Result

    pos = InStr(Result, Chr$(0))
    Result = Left$(Result, pos - 1)

    FindFileExecutable = Result
    
End Function

' Create the complete directory path.
Private Sub MakePath(ByVal Path As String)
    
    Dim directories As Variant
    Dim i As Integer
    Dim new_dir As String
    Dim dir_path As String

    ' Break the path into directories.
    directories = Split(Path, "\")

    ' Build the subdirectories.
    For i = LBound(directories) To UBound(directories)
        ' Get the next directory in the path.
        new_dir = directories(i)

        ' Make sure the entry is non-empty.
        If Len(new_dir) > 0 Then
            ' Add the new directory to the path.
            dir_path = dir_path & new_dir & "\"

            ' Make sure we don't just have a drive
            ' specification.
            If Right$(new_dir, 1) <> ":" Then
                ' See if the directory already exists.
                If Dir$(dir_path, vbDirectory) = "" Then
                    ' The directory doesn't exist.
                    ' Make it.
                    MkDir dir_path
                End If
            End If
        End If
    Next i
    
End Sub

Private Function GetFileName(s As String) As String
    Dim N As Integer
    
    GetFileName = ""
    For N = Len(s) To 1 Step -1
        If Mid(s, N, 1) = "\" Or Mid(s, N, 1) = "/" Then Exit For
    Next N
    GetFileName = Right(s, Len(s) - N)
End Function

'If Right(Ftp1.Directory, 1) = "/" Then
'    m_Item.Tag = Ftp1.Directory + s
'Else
'    m_Item.Tag = Ftp1.Directory + "/" + s
'End If

Private Sub SetFileIcon(Item As ListItem)
    ' Set some standard icon types
    Select Case UCase(Right(Item.Text, 4))
        Case ".TXT"
            Item.SmallIcon = "TextFile"
        Case ".ZIP", ".ARJ", ".TGZ", ".TAR", ".LZH", ".LZA"
            Item.SmallIcon = "ZipFile"
        Case ".EXE", ".COM", ".BAT"
            Item.SmallIcon = "ExeFile"
        Case Else
            Item.SmallIcon = "File"
    End Select
End Sub

Public Function FileExists(ByVal strFilePath As String) As Boolean
    
    ' Validate parameters
    strFilePath = Trim(strFilePath)
    If strFilePath = "" Then Exit Function
    
    ' Check if the file exists (reguardless of it's attributes)
    If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
        If (GetAttr(strFilePath) And vbDirectory) = vbDirectory Then
            FileExists = False
        Else
            FileExists = True
        End If
    Else
          FileExists = False
    End If
    
End Function

Public Function GetAppVersion() As String
    On Error Resume Next
    GetAppVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function

'====================================================================================
' EndTheProgram
'
' Description:
' ------------
' Unloads all open forms and ends the program.  The ExitingProgram variable is set
' to TRUE to let the rest of the program know that it's trying to end the program.
'
' Param                 Description
' ---------------------------------
' CallEnd               Optional. If set to TRUE, the "End" function is called
'                       to make sure the program is ended.  If the program being
'                       ended contains any kind of sub-classing, this should be
'                       set to FALSE to avoid the program crashing.
'
' Return:
' -------
' Program ends
'
'====================================================================================
Public Sub EndTheProgram(Optional ByVal CallEnd As Boolean = False)
On Error Resume Next
  
  Dim Form As Form
  
  'ExitingProgram = True
  
  For Each Form In Forms
    Unload Form
    Set Form = Nothing
  Next
  
  If CallEnd = True Then
    End
  End If
  
End Sub

'====================================================================================
' FileInUse
'
' Description:
' ------------
' Checks to see if the specified file is currently in use.
'
' Param                 Description
' ---------------------------------
' strFilePath           Fully qualified path to the file to check.
'
' Return:
' -------
' TRUE  = Succeeded
' FALSE = Failed
'
'====================================================================================
Public Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  ' Validate parameters
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  ' Attempt to open the file EXCLUSIVELY... if this fails, another process is using the file
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Public Function CopyFile(ByVal sSourceFile As String, _
        ByVal sDestinationFile As String, _
        Optional ByVal bFailIfDestExists As Boolean = False) As Boolean
    
    CopyFile = CBool(CopyFileAPI(sSourceFile, sDestinationFile, Abs(CLng(bFailIfDestExists))))

End Function

'Public Function EncryptText(ByVal sText As String) As String
'    RC4Ini EncryptionPassword
'    EncryptText = EncodeStr64(EnDeCrypt(sText))
'End Function
'
'Public Function DecryptText(ByVal sText As String) As String
'    RC4Ini EncryptionPassword
'    DecryptText = EnDeCrypt(DecodeStr64(sText))
'End Function
'
'Return the directory with a "\" or "/" depend on the system.
Public Function ReturnPath(ByVal sPath As String, Optional ByVal blnSlashUnix = False) As String
    
    Dim Path As String
    Dim first As Long
    Dim last As Long
    
    On Error GoTo ErrHandler
    Path = sPath
    
    first = InStr(1, Path, "[")
    last = InStr(1, Path, "]")
    If first > 0 And last > 0 Then
        Path = Trim(Mid(Path, 1, first - 1) & Mid(Path, last + 1))
    Else
        Path = Trim(Path)
    End If
    
    If blnSlashUnix Then
        If Right(Path, 1) <> "/" Then
            Path = Path & "/"
        End If
    Else
        If Right(Path, 1) <> "\" Then
            Path = Path & "\"
        End If
    End If
    
    ReturnPath = Path
    Exit Function
    
ErrHandler:
    ReturnPath = sPath
End Function

Public Function GetDriveSerialNumber(strDrive As String) As Long
    'This function returns the serial number of the Hard-Disk
    Dim SerialNum As Long, Res As Long
    Dim temp1 As String, Temp2 As String
    temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    Res = GetVolumeInformation(strDrive, temp1, Len(temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetDriveSerialNumber = SerialNum
End Function

' Return the Windows directory's drive letter.
Private Function SystemDrive() As String
    SystemDrive = Left$(WindowsDirectory(), 1)
End Function

' Return the Windows directory.
Private Function WindowsDirectory() As String
    Dim windows_dir As String
    Dim length As Long

    ' Get the Windows directory.
    windows_dir = Space$(MAX_PATH)
    length = GetWindowsDirectory(windows_dir, Len(windows_dir))
    WindowsDirectory = Left$(windows_dir, length)
End Function

Public Function GetSettings(sAppName, _
                        sSection, _
                        sKey, _
                        Optional sDefault = vbNullString)
    
    'GetSetting(appname, section, key[, default])

    Dim cR As New cRegistry
    cR.ClassKey = HKEY_LOCAL_MACHINE
    'cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    'cR.SectionKey = "Software\Mewsoft\Backupawy\Settings"
    cR.SectionKey = "Software\" & sAppName & "\" & sSection
    cR.ValueKey = sKey
    cR.Default = sDefault
    cR.ValueType = REG_SZ
    GetSettings = cR.Value

End Function

Public Sub SaveSettings(sAppName As String, _
                        sSection As String, _
                        sKey As String, sValue As Variant)
    
    'GetSetting(appname, section, key[, default])
    
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_LOCAL_MACHINE
    'cR.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    'cR.SectionKey = "Software\Mewsoft\Backupawy\Settings"
    cR.SectionKey = "Software\" & sAppName & "\" & sSection
    cR.ValueKey = sKey
    'cR.Default = sDefault
    cR.ValueType = REG_SZ
    cR.Value = sValue

End Sub

'====================================================================
Public Function DirectoryFiles(ByVal DirPath As String) As String()
'***************************************************
'PURPOSE: RETURN AN ARRAY CONTAINING NAME OF ALL FILES IN
'DIRECTORY SPECIFIED BY DIR PATH

'PARAMETER:

  'DIRPATH: A VALID DRIVE OR SUBDIRECTORY ON YOUR SYSTEM,
  'ENDING WITH FORWARD SLASH (\) CHARACTER, OR
  'A DRIVE OR SUBDIRECTORY FOLLOWED BY A WILD CARD
  'STRING (e.g., C:\WINDOWS\*.txt)

'RETURNS: A STRING ARRAY WITH THE NAMES OF ALL FILENAMES
'IN THE DIRECTORY, INCLUDING HIDDEN, SYSTEM, AND READ-ONLY FILES
'THE FUNCTION IS NON RECURSIVE, I.E., IT DOES NOT SEARCH
'SUBDIRECTORIES UNDERNEATH DIRPATH

'REQUIRES: VB6, BECAUSE IT RETURNS A STRING ARRAY

'EXAMPLE
'Dim sFiles() As String
'Dim lCtr As Long

'sFiles = AllFiles("C:\windows\")
'For lCtr = 0 To UBound(sFiles)
'    Debug.Print sFiles(lCtr)
'Next
'********************************************************
    Dim sFile As String
    Dim lElement As Long
    Dim sAns() As String
    
    ReDim sAns(0) As String
    
    sAns(0) = ""
    
    sFile = Dir(DirPath, vbNormal + vbHidden + vbReadOnly + _
       vbSystem + vbArchive)
       
    If sFile <> "" Then
        sAns(0) = sFile
        Do
            sFile = Dir
            If sFile = "" Then Exit Do
            lElement = IIf(sAns(0) = "", 0, UBound(sAns) + 1)
            ReDim Preserve sAns(lElement) As String
            sAns(lElement) = sFile
        Loop
    End If
    DirectoryFiles = sAns
End Function

Public Function APIAllFiles(ByVal DirPath As String) As String()
'***************************************************
'SEE COMMENTS FOR ALLFILES.  PURPOSE AND INSTRUCTIONS ARE
'EXACTLY THE SAME, EXCEPT THIS FILE USES THE WIN32 API
'***************************************************

Dim sFile As String
Dim lElement As Long
Dim sAns() As String
Dim lFirstRet As Long, lNextRet
Dim typFindData As WIN32_FIND_DATA
Dim sTemp As String
Dim lAttr As Long

ReDim sAns(0) As String

If Right(DirPath, 1) = "\" Then DirPath = DirPath & "*.*"

'Get First File
lFirstRet = FindFirstFile(DirPath, typFindData)
If lFirstRet <> -1 Then
    lAttr = typFindData.dwFileAttributes

    'Check if this is a directory.  This is probably slowing down
    'the function.  If you know you won't have subdirectories,
    'or you want to include directories, delete this check

    If Not isDirectory(lAttr) Then

        'strip null terminator
       sAns(0) = StripNull(typFindData.cFileName)
    End If
    
    'Continue searching until all files in directory are found
    Do
        lNextRet = FindNextFile(lFirstRet, typFindData)
        If lNextRet = ERROR_NO_MORE_FILES Or _
           lNextRet = 0 Then Exit Do
        lAttr = typFindData.dwFileAttributes
            'Again, check if its a subdirectory
            If Not isDirectory(lAttr) Then
                lElement = IIf(sAns(0) = "", 0, UBound(sAns) + 1)
                ReDim Preserve sAns(lElement) As String
                sAns(lElement) = StripNull(typFindData.cFileName)
            End If
    Loop
End If

FindClose lFirstRet
APIAllFiles = sAns

End Function

Private Function StripNull(ByVal InString As String) As String

'Input: String containing null terminator (Chr(0))
'Returns: all character before the null terminator

Dim iNull As Integer
If Len(InString) > 0 Then
    iNull = InStr(InString, vbNullChar)
    Select Case iNull
    Case 0
        StripNull = InString
    Case 1
        StripNull = ""
    Case Else
       StripNull = Left$(InString, iNull - 1)
   End Select
End If

End Function

Public Function isDirectory(FileAttr As Long) As Boolean

Dim bAns As Boolean
Dim lDir As Long
Dim lHidden As Long
Dim lSystem As Long
Dim lReadOnly As Long

lDir = FILE_ATTRIBUTE_DIRECTORY
lHidden = FILE_ATTRIBUTE_HIDDEN
lSystem = FILE_ATTRIBUTE_SYSTEM
lReadOnly = FILE_ATTRIBUTE_READONLY
    
isDirectory = FileAttr = lDir Or FileAttr = _
    lDir + lHidden Or FileAttr = lDir + lSystem _
    Or FileAttr = lDir + lReadOnly Or FileAttr = _
    lDir + lHidden + lSystem Or FileAttr = _
    lDir + lHidden + lReadOnly Or FileAttr = _
    lDir + lSystem + lReadOnly Or _
    FileAttr = lDir + lSystem + lHidden + lReadOnly

End Function

'====================================================================
Public Function SafeLBound(Niz() As String) As Long
    On Error GoTo SafeLBound_Err
    
    SafeLBound = LBound(Niz)
SafeLBound_Exit:
    Exit Function
SafeLBound_Err:
    If Err.Number = 9 Then SafeLBound = -1
    Resume SafeLBound_Exit
End Function

Public Function SafeUBound(Niz() As String) As Long
    On Error GoTo SafeUBound_Err
    
    SafeUBound = UBound(Niz)
SafeUBound_Exit:
    Exit Function
SafeUBound_Err:
    If Err.Number = 9 Then SafeUBound = -1
    Resume SafeUBound_Exit
End Function

Public Sub GetRGB(Color As Long, R As Long, G As Long, b As Long)
    Dim HexColor As String
        
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    
    R = "&H" & Mid(HexColor, 5, 2)
    G = "&H" & Mid(HexColor, 3, 2)
    b = "&H" & Mid(HexColor, 1, 2)
End Sub

'Function to get the total system RAM and available RAM
Public Sub GetMemoryStatistics(ByRef TotalRAM As Long, ByRef AvailableRAM As Long)
   On Error GoTo ErrHandler
   Dim Mem As MEMORYSTATUS
   Call GlobalMemoryStatus(Mem)
   'return megabytes in MB
   TotalRAM = (Mem.dwTotalPhys / 1048576)
   AvailableRAM = (Mem.dwAvailPhys / 1048576)
   Exit Sub
ErrHandler:
   Debug.Print "Error: " + Err.Description
End Sub

Public Function GetMemoryAvailable() As Long
   On Error GoTo ErrHandler
   Dim Mem As MEMORYSTATUS
   Call GlobalMemoryStatus(Mem)
   'TotalRAM = (Mem.dwTotalPhys / 1048576)
   GetMemoryAvailable = (Mem.dwAvailPhys / 1048576)
   Exit Function
ErrHandler:
   Debug.Print "Error: " + Err.Description
End Function


