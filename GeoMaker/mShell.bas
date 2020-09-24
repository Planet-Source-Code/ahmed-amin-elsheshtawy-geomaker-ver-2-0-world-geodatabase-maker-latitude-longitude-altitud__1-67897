Attribute VB_Name = "mShell"
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

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

Public Function ShellDocument(sDocName As String, _
                    Optional ByVal Action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState) As Boolean
    Dim Response
    Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
End Function

Public Sub OpenFolder(sFolderPath As String)
  
    Const SW_SHOWNORMAL As Long = 1
    Const SW_SHOWMAXIMIZED As Long = 3
    Const SW_SHOWDEFAULT As Long = 10
  
    On Error Resume Next
    RunShellExecute "Open", "explorer.exe", "/e, " + sFolderPath, vbNullString, SW_SHOWDEFAULT
End Sub

Private Sub RunShellExecute(sTopic As String, _
                            sFile As Variant, _
                            sParams As Variant, _
                            sDirectory As Variant, _
                            nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

End Sub

