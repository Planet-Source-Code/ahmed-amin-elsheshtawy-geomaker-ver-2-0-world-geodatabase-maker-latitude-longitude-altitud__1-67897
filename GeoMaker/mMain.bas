Attribute VB_Name = "mMain"
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

Public fMainForm As frmMain

'====================================================================
'====================================================================
'Global Declaration
Global Const gAppName = "GeoMaker"
Public Const AppRegPath = "Mewsoft\GeoMaker"
Public Const AppRegSettingsSection = "Settings"

Public ThreadsCount As Long

'====================================================================

'====================================================================
'====================================================================
Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub
'====================================================================
'====================================================================

