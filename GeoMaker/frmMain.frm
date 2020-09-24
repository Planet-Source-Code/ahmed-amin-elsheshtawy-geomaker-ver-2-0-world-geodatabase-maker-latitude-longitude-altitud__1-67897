VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{49CF2586-6B7A-41FF-96BF-6D26C500CC8A}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GeoMaker"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   4560
      Top             =   3480
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   330
      Left            =   5580
      TabIndex        =   11
      Top             =   0
      Width           =   375
      ExtentX         =   661
      ExtentY         =   582
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet Inet 
      Index           =   0
      Left            =   9315
      Top             =   6975
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   11295
      Top             =   6795
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6900
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   12171
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   2646
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Database"
      TabPicture(0)   =   "frmMain.frx":113A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Settings"
      TabPicture(1)   =   "frmMain.frx":1156
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "frmMain.frx":1172
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   6360
         Left            =   180
         TabIndex        =   14
         Top             =   405
         Width           =   12300
         Begin VB.Label Label15 
            Caption         =   "Author: Dr. Ahmed Amin Elsheshtawy, Ph.D."
            Height          =   255
            Left            =   720
            TabIndex        =   44
            Top             =   1680
            Width           =   4815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   720
            X2              =   6885
            Y1              =   3195
            Y2              =   3195
         End
         Begin VB.Label lblDisclaimer 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":118E
            ForeColor       =   &H00000000&
            Height          =   900
            Left            =   720
            TabIndex        =   38
            Tag             =   "Warning: ..."
            Top             =   3330
            Width           =   6210
         End
         Begin VB.Label lblVersion 
            BackStyle       =   0  'Transparent
            Caption         =   "Version"
            Height          =   225
            Left            =   720
            TabIndex        =   37
            Tag             =   "Version"
            Top             =   960
            Width           =   4725
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            Index           =   1
            X1              =   720
            X2              =   6885
            Y1              =   3195
            Y2              =   3195
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "GeoMaker"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   720
            TabIndex        =   36
            Tag             =   "Application Title"
            Top             =   540
            Width           =   4725
         End
         Begin VB.Label lblSalesEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "sales@mewsoft.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2220
            MouseIcon       =   "frmMain.frx":12B2
            MousePointer    =   99  'Custom
            TabIndex        =   35
            Top             =   2775
            Width           =   2580
         End
         Begin VB.Label lblSupportEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "support@mewsoft.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2235
            MouseIcon       =   "frmMain.frx":15BC
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   2415
            Width           =   2580
         End
         Begin VB.Label lblMewsoft 
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.mewsoft.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2220
            MouseIcon       =   "frmMain.frx":18C6
            MousePointer    =   99  'Custom
            TabIndex        =   33
            Top             =   2115
            Width           =   2550
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Emails:"
            Height          =   255
            Left            =   720
            TabIndex        =   32
            Top             =   2745
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Support Email:"
            Height          =   255
            Left            =   720
            TabIndex        =   31
            Top             =   2415
            Width           =   1155
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Website address:"
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Top             =   2115
            Width           =   1275
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyrights (c) Mewsoft Corporation. All rights reserved."
            Height          =   255
            Left            =   720
            TabIndex        =   29
            Top             =   1320
            Width           =   4695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6315
         Left            =   -74820
         TabIndex        =   13
         Top             =   405
         Width           =   12300
         Begin VB.CommandButton cmdOpenDataFolder 
            Caption         =   "Open"
            Height          =   285
            Left            =   8775
            TabIndex        =   43
            Top             =   4860
            Width           =   825
         End
         Begin VB.TextBox txtDataFolder 
            Height          =   285
            Left            =   360
            TabIndex        =   42
            Top             =   4860
            Width           =   7395
         End
         Begin VB.CommandButton cmdDataFolderBrowse 
            Caption         =   "Browse"
            Height          =   285
            Left            =   7785
            TabIndex        =   40
            Top             =   4860
            Width           =   870
         End
         Begin VB.CommandButton cmdResetSettings 
            Caption         =   "Reset"
            Height          =   285
            Left            =   6975
            TabIndex        =   28
            Top             =   5625
            Width           =   1500
         End
         Begin VB.CommandButton cmdSaveSettings 
            Caption         =   "Save Changes"
            Height          =   285
            Left            =   4995
            TabIndex        =   27
            Top             =   5625
            Width           =   1725
         End
         Begin VB.CommandButton cmdTabFormat 
            Caption         =   "Tab Format"
            Height          =   285
            Left            =   8370
            TabIndex        =   26
            Top             =   855
            Width           =   1365
         End
         Begin VB.CommandButton cmdCSVFormat 
            Caption         =   "CSV Format"
            Height          =   285
            Left            =   6885
            TabIndex        =   25
            Top             =   855
            Width           =   1365
         End
         Begin VB.CommandButton cmdDefaultFormat 
            Caption         =   "Default Format"
            Height          =   285
            Left            =   5175
            TabIndex        =   24
            Top             =   855
            Width           =   1545
         End
         Begin VB.TextBox txtFormat 
            Height          =   285
            Left            =   315
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   1170
            Width           =   11805
         End
         Begin VB.Label Label14 
            Caption         =   "Output data folder:"
            Height          =   240
            Left            =   270
            TabIndex        =   41
            Top             =   4590
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Tab"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   630
            TabIndex        =   39
            Top             =   3870
            Width           =   555
         End
         Begin VB.Label Label8 
            Caption         =   "%Altitude%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   630
            TabIndex        =   23
            Top             =   3555
            Width           =   1320
         End
         Begin VB.Label Label7 
            Caption         =   "%Latitude%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   630
            TabIndex        =   22
            Top             =   2925
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "%Longitude%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   630
            TabIndex        =   21
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "%County%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   630
            TabIndex        =   20
            Top             =   2610
            Width           =   1140
         End
         Begin VB.Label Label4 
            Caption         =   "%Region%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   630
            TabIndex        =   19
            Top             =   2340
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "%Country%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   630
            TabIndex        =   18
            Top             =   2025
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "Use these to format each line in the output data file (Click to insert):"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   1665
            Width           =   5820
         End
         Begin VB.Label Label1 
            Caption         =   "Output file format:"
            Height          =   240
            Left            =   270
            TabIndex        =   15
            Top             =   765
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6405
         Left            =   -74820
         TabIndex        =   2
         Top             =   405
         Width           =   12345
         Begin VB.CommandButton cmdClearData 
            Caption         =   "Clear Data"
            Height          =   285
            Left            =   9540
            TabIndex        =   12
            Top             =   270
            Width           =   1095
         End
         Begin VB.CommandButton cmdPause 
            Caption         =   "Pause"
            Height          =   285
            Left            =   5895
            TabIndex        =   10
            Top             =   270
            Width           =   1275
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Height          =   285
            Left            =   7245
            TabIndex        =   9
            Top             =   270
            Width           =   1230
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start"
            Height          =   285
            Left            =   4545
            TabIndex        =   8
            Top             =   270
            Width           =   1275
         End
         Begin MSComctlLib.ListView lvwData 
            Height          =   5640
            Left            =   3510
            TabIndex        =   7
            Top             =   630
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   9948
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvwCountries 
            Height          =   5640
            Left            =   135
            TabIndex        =   6
            Top             =   630
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   9948
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CommandButton cmdDeSelectAllCountries 
            Caption         =   "Select None"
            Height          =   285
            Left            =   2250
            TabIndex        =   5
            Top             =   270
            Width           =   1140
         End
         Begin VB.CommandButton cmdSelectAllCountries 
            Caption         =   "Select All"
            Height          =   285
            Left            =   1215
            TabIndex        =   4
            Top             =   270
            Width           =   1005
         End
         Begin VB.CommandButton cmdRefreshCountries 
            Caption         =   "Refresh"
            Height          =   285
            Left            =   180
            TabIndex        =   3
            Top             =   270
            Width           =   1005
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3581
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim CountryName As Collection
Dim bStopRequest As Boolean
Dim bPauseRequest As Boolean
Dim CurrentCountryFileNum As Long
Dim TotalFound As Long
Dim CountryCities As Long
Dim TotalQueries As Long
Dim SelectedCountries As Long
Dim CountriesDone As Long
Dim OutputFormat As String
Const DefaultFormat = "%Country%|%Region%|%County%|%Latitude%|%Longitude%|%Altitude%"
Dim OutputFolder As String

'====================================================================
'====================================================================

Private Sub Form_Load()
    '
    'http://heavens-above.com/countries.aspx
    '----------------------------------------------------------------
    Me.Left = GetSettings(AppRegPath, "Settings", "MainLeft", 1000)
    Me.Top = GetSettings(AppRegPath, "Settings", "MainTop", 1000)
    'Me.Width = GetSettings(AppRegPath, "Settings", "MainWidth", 11685)
    'Me.Height = GetSettings(AppRegPath, "Settings", "MainHeight", 8085)
    '----------------------------------------------------------------
    OutputFormat = GetSettings(AppRegPath, "Settings", "OutputFormat", DefaultFormat)
    txtFormat.Text = OutputFormat
    OutputFolder = GetSettings(AppRegPath, "Settings", "OutputFolder", ReturnPath(AppPath + "data"))
    txtDataFolder.Text = OutputFolder
    CreateDirectoryStruct OutputFolder
    '----------------------------------------------------------------
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdClearData.Enabled = False
    '------------------------------------------------------
    bStopRequest = False
    bPauseRequest = False
    '------------------------------------------------------
    WebBrowser1.Navigate2 "about:blank"
    WebBrowser1.ZOrder 1
    WebBrowser1.Move -1000, -1000
    '------------------------------------------------------
    lvwCountries.ColumnHeaders.Add , , "Country", lvwCountries.Width
    LoadCountries
    '------------------------------------------------------
    lvwData.ColumnHeaders.Add , , "#", 1000
    lvwData.ColumnHeaders.Add , , "Country", 1500
    lvwData.ColumnHeaders.Add , , "Region (State)", 1500
    lvwData.ColumnHeaders.Add , , "County (City)", 1500
    lvwData.ColumnHeaders.Add , , "Latitude", 1000
    lvwData.ColumnHeaders.Add , , "Longitude", 1000
    lvwData.ColumnHeaders.Add , , "Altitude", 1000 'elevation
    '------------------------------------------------------
    ResetStats
    '------------------------------------------------------
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    '------------------------------------------------------
    '----------------------------------------------------------------
    WindowsXPC1.ColorScheme = XP_Blue
    WindowsXPC1.InitSubClassing
    '----------------------------------------------------------------
    SSTab1.Tab = 0
    
End Sub

Private Sub cmdCSVFormat_Click()
    txtFormat.Text = """%Country%"",""%Region%"",""%County%"",""%Latitude%"",""%Longitude%"",""%Altitude%"""
End Sub

Private Sub cmdDefaultFormat_Click()
    txtFormat.Text = DefaultFormat
End Sub

Private Sub cmdResetSettings_Click()
    txtFormat.Text = DefaultFormat
End Sub

Private Sub cmdTabFormat_Click()
    txtFormat.Text = "%Country%" + vbTab + "%Region%" + vbTab + "%County%" + vbTab + "%Latitude%" + vbTab + "%Longitude%" + vbTab + "%Altitude%"
End Sub

Private Sub CheckPauseRequest()
    
    Do While bPauseRequest = True
        If bStopRequest = True Then Exit Do
        DoEvents
    Loop
    
End Sub

Private Sub cmdPause_Click()

    bPauseRequest = Not bPauseRequest
    If bPauseRequest = True Then
        cmdPause.Caption = "Resume"
        UpdateStatus "Pausing"
    Else
        cmdPause.Caption = "Pause"
        UpdateStatus "Runing"
    End If
    
End Sub

Private Sub cmdStop_Click()
    UpdateStatus "Stopping"
    bStopRequest = True
End Sub

Private Sub cmdClearData_Click()
    ResetStats
End Sub

Private Sub ResetStats()
    
    lvwData.ListItems.Clear
    
    TotalFound = 0
    CountryCities = 0
    TotalQueries = 0
    CountriesDone = 0
    SelectedCountries = 0

    UpdateStatus "Ready"
    UpdateCities 0
    UpdateTotalFound 0
    UpdateCountry "-"
    UpdateQuery "-"
    UpdateQueries 0
    UpdateCountriesDone 0
    
    DoEvents
End Sub

Private Sub cmdStart_Click()

    Dim X As Long
    Dim CityCode As Long
    Dim CountryCode As String
    Dim CountryFile As String
    
    cmdStart.Enabled = False
    cmdPause.Enabled = True
    cmdStop.Enabled = True
    cmdClearData.Enabled = False
    
    bStopRequest = False
    bPauseRequest = False
    
    lvwData.ListItems.Clear
    TotalFound = 0
    CountryCities = 0
    TotalQueries = 0
    CountriesDone = 0
    
    UpdateStatus "Runing"
    '----------------------------------------------------------------
    CreateDirectoryStruct OutputFolder
    '----------------------------------------------------------------
    SelectedCountries = 0
    For X = 1 To lvwCountries.ListItems.Count
        If lvwCountries.ListItems.Item(X).Checked = True Then
            SelectedCountries = SelectedCountries + 1
        End If
    Next
    
    UpdateCountriesDone CountriesDone
    
    'search for all cities in each selected country
    'send a search query for all cities starting from a* to z*
    For X = 1 To lvwCountries.ListItems.Count
        If lvwCountries.ListItems.Item(X).Checked = True Then
            CountryCode = lvwCountries.ListItems.Item(X).Key
            '------------------------------------------------------------
            UpdateCountry CountryCode
            
            CountryCities = 0
            'create the country data file
            CountryFile = AppPath + "data\" + CountryCode + ".txt"
            If Dir(CountryFile + ".bak") <> "" Then
                Kill CountryFile + ".bak"
            End If
            If Dir(CountryFile) <> "" Then
                Name CountryFile As CountryFile + ".bak"
            End If
            
            CurrentCountryFileNum = FreeFile
            Open CountryFile For Output As CurrentCountryFileNum
            '------------------------------------------------------------
            For CityCode = Asc("a") To Asc("z")
                'search query from a* to z*
                If bStopRequest = True Then
                    Close #CurrentCountryFileNum
                    GoTo StopRequested
                End If
                CheckPauseRequest
                UpdateQuery Chr(CityCode) + "*"
                MakeGeoData CountryCode, Chr(CityCode) + "*"
            Next
            '------------------------------------------------------------
            Close #CurrentCountryFileNum
            CountriesDone = CountriesDone + 1
            UpdateCountriesDone CountriesDone
        End If
    Next

StopRequested:

    bStopRequest = False
    bPauseRequest = False
    
    cmdStart.Enabled = True
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdClearData.Enabled = True
    
    UpdateStatus "Ready"
    
End Sub

Sub MakeGeoData(ByVal CountryCode As String, SearchQuery As String)
    
    Dim URL As String, HtmlText As String
    Dim Pattern As String
    Dim MyRegExp As RegExp
    Dim Matches As MatchCollection
    Dim Match As Match
    Dim itmX As ListItem
    Dim County, State, Latitude, Longitude, Elevation, Alias
    Dim colTables
    Dim TR As HTMLTableRow
    Dim td As IHTMLTableCell
    Dim tblTable As IHTMLTable
    Dim CellCollection As IHTMLElementCollection
    Dim RowCollection As IHTMLElementCollection
    Dim colCurrentTable As IHTMLTable
    Dim X As Long, Y As Long
    Dim FileNum As Long
    Dim Counter As Long
    Dim CityCode As Long
    Dim NewSearchQuery As String
    Dim OutputLine As String
    
    If CountryCode = "" Then Exit Sub
    
    URL = "http://www.heavens-above.com/selecttown.asp?CountryID="
    URL = URL + CountryCode + "&loc=Unspecified"
    '------------------------------------------------------
    ' Navigate to the select country page:
    ' http://heavens-above.com/selecttown.asp?CountryID=AF
    WebBrowser1.Navigate URL
    ' Wait until the browser finishs loading the page completely
    Do: DoEvents: Loop While WebBrowser1.Busy
    Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE: DoEvents: Loop
    '------------------------------------------------------
    If bStopRequest = True Then Exit Sub
    CheckPauseRequest
    '------------------------------------------------------
    ' Fill the search form field and click the submit button
    WebBrowser1.Document.Forms(0).elements("Search").Value = SearchQuery
    WebBrowser1.Document.Forms(0).elements("CountryID").Value = CountryCode
    WebBrowser1.Document.Forms(0).elements(2).Click
    ' Wait until the browser finishs loading the page completely
    Do: DoEvents: Loop While WebBrowser1.Busy
    Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE: DoEvents: Loop
    '------------------------------------------------------
    If bStopRequest = True Then Exit Sub
    CheckPauseRequest
    TotalQueries = TotalQueries + 1
    UpdateQueries TotalQueries
    '------------------------------------------------------
    ' get the full html page source text
    HtmlText = WebBrowser1.Document.body.innerHTML
    'Set colTables = WebBrowser1.Document.All.tags("TABLE")
    
    'this is the data table on the page
    Set tblTable = WebBrowser1.Document.All.tags("TABLE")(2)
    '------------------------------------------------------------
    'The first row is the table header : Name  Region  Latitude  Longitude  Elevation
    ' starting from the second row is the data
    ' loop throgh all the table rows and get each cell text and parse the info
    Set itmX = Nothing
    For X = 1 To tblTable.rows.length - 1
        '------------------------------------------------------
        If bStopRequest = True Then Exit Sub
        CheckPauseRequest
        '------------------------------------------------------
        Set TR = tblTable.rows(X)
        
        If TR.cells.length >= 6 Then
            County = Trim(TR.cells(0).innerText)
            State = Trim(TR.cells(1).innerText)
            Latitude = CDbl(Val(TR.cells(2).innerText))
            Longitude = CDbl(Val(TR.cells(3).innerText))
            Elevation = CDbl(Val(TR.cells(4).innerText))
        ElseIf TR.cells.length = 5 Then
            County = Trim(TR.cells(0).innerText)
            'State = Trim(TR.cells(1).innerText)
            Latitude = CDbl(Val(TR.cells(1).innerText))
            Longitude = CDbl(Val(TR.cells(2).innerText))
            Elevation = CDbl(Val(TR.cells(3).innerText))
        Else
            Exit Sub
        End If
        
        County = Replace(County, vbCrLf, "")
        County = Replace(County, vbCr, "")
        County = Replace(County, vbLf, "")
        
        'Some cities has an alias name so their cell entry in the form:
        '       Cabal
        '       (alias for Chabal)
        ' you will find the city name then line break then (alias for ...)
        Y = InStr(1, County, "alias for")
        If Y > 0 Then
            Alias = County
            County = Trim(Left(County, Y - 2))
            Alias = Right(Alias, Len(Alias) - Y - Len("alias for"))
            Alias = Replace(Alias, ")", "")
            Alias = Trim(Alias)
        End If
        
        County = Trim(County)
        '------------------------------------------------------------
        ' display the data line in the listview
        Counter = lvwData.ListItems.Count + 1
        Set itmX = lvwData.ListItems.Add(, , Counter)
        
        itmX.SubItems(1) = CountryName(CountryCode)
        itmX.SubItems(2) = State
        itmX.SubItems(3) = County
        itmX.SubItems(4) = Latitude
        itmX.SubItems(5) = Longitude
        itmX.SubItems(6) = Elevation
        If (X Mod 40) = 0 Then
            'DoEvents
            'itmX.EnsureVisible
        End If
        '------------------------------------------------------------
        TotalFound = TotalFound + 1
        CountryCities = CountryCities + 1
        UpdateCities CountryCities
        UpdateTotalFound TotalFound
        '------------------------------------------------------------
        'save the data line to the country data file
        OutputLine = OutputFormat
        OutputLine = Replace(OutputLine, "%Country%", CountryCode)
        OutputLine = Replace(OutputLine, "%Region%", State)
        OutputLine = Replace(OutputLine, "%County%", County)
        OutputLine = Replace(OutputLine, "%Latitude%", Latitude)
        OutputLine = Replace(OutputLine, "%Longitude%", Longitude)
        OutputLine = Replace(OutputLine, "%Altitude%", Elevation)
        Print #CurrentCountryFileNum, OutputLine
        'Print #CurrentCountryFileNum, CountryCode + "|" + State + "|" + County + "|" + _
        '                CStr(Latitude) + "|" + CStr(Longitude) + "|" + CStr(Elevation)
        '------------------------------------------------------------
    Next
    If lvwData.ListItems.Count > 0 Then
        Set itmX = lvwData.ListItems.Item(lvwData.ListItems.Count)
        If Not itmX Is Nothing Then
            itmX.EnsureVisible
            DoEvents
        End If
    End If
    '----------------------------------------------------------------
    If InStr(1, HtmlText, "cut-off after 200 towns") > 1 Then
        '------------------------------------------------------------
        For CityCode = Asc("a") To Asc("z")
            CheckPauseRequest
            If bStopRequest = True Then GoTo StopRequested
            'search query from SearchQuery+a* to SearchQuery+z*
            NewSearchQuery = SearchQuery
            NewSearchQuery = Replace(NewSearchQuery, "*", "")
            NewSearchQuery = NewSearchQuery + Chr(CityCode) + "*"
            UpdateQuery NewSearchQuery
            MakeGeoData CountryCode, NewSearchQuery
        Next
        '------------------------------------------------------------
    End If
    
StopRequested:

End Sub

Sub UpdateCities(CountryCities As Long)
    StatusBar.Panels(4).Text = "Cities: " + CStr(CountryCities)
End Sub

Sub UpdateTotalFound(TotalFound As Long)
    StatusBar.Panels(5).Text = "All Cities: " + CStr(TotalFound)
End Sub

Sub UpdateCountry(CountryCode As String)
    On Error Resume Next
    If CountryCode <> "-" And CountryCode <> "" Then
        StatusBar.Panels(2).Text = "Country: " + CountryName(CountryCode)
    Else
        StatusBar.Panels(2).Text = "Country: " + "-"
    End If
End Sub
            
Sub UpdateStatus(StatusText As String)
    StatusBar.Panels(1).Text = "Status: " + StatusText
End Sub
            
Sub UpdateQuery(SearchQuery As String)
    StatusBar.Panels(3).Text = "Query: " + SearchQuery
End Sub
            
Sub UpdateQueries(Queries As Long)
    StatusBar.Panels(6).Text = "Queries: " + CStr(Queries)
End Sub
            
Sub UpdateCountriesDone(Countries As Long)
    StatusBar.Panels(7).Text = "Countries: " + CStr(Countries) + "/" + CStr(SelectedCountries)
End Sub
            
Private Sub cmdSelectAllCountries_Click()
    Dim X As Long
    For X = 1 To lvwCountries.ListItems.Count
        lvwCountries.ListItems.Item(X).Selected = False
        lvwCountries.ListItems.Item(X).Checked = True
    Next
End Sub

Private Sub cmdDeSelectAllCountries_Click()
    Dim X As Long
    For X = 1 To lvwCountries.ListItems.Count
        lvwCountries.ListItems.Item(X).Selected = False
        lvwCountries.ListItems.Item(X).Checked = False
    Next
End Sub

Private Sub cmdRefreshCountries_Click()
    
    Me.MousePointer = vbHourglass
    GetWorldIndex
    Me.MousePointer = vbDefault
    
    UpdateStatus "Ready"
End Sub

Sub LoadCountries()
    
    Dim itmX As ListItem
    Dim FileNum As Long
    Dim TextLine As String
    Dim Parts() As String
    
    lvwCountries.ListItems.Clear
     
    FileNum = FreeFile
    Open AppPath + "Countries.txt" For Input As FileNum
    
    Set CountryName = New Collection
    
    While Not EOF(FileNum)
        Line Input #FileNum, TextLine
        Parts = Split(TextLine, "=")
        ReDim Preserve Parts(2) As String
        If Parts(0) <> "" And Parts(1) <> "" Then
            Set itmX = lvwCountries.ListItems.Add(, Parts(0), Parts(1))
        End If
        CountryName.Add Item:=Parts(1), Key:=Parts(0)

    Wend
    Close #FileNum
    
End Sub

Sub GetWorldIndex()

    Dim MyRegExp As RegExp
    Dim Matches As MatchCollection
    Dim Match As Match
    Dim HtmlText As String
    Dim URL As String
    Dim CountryName As String
    Dim CountryCode As String
    Dim itmX As ListItem
    Dim FileNum As Long
        
    
    URL = "http://heavens-above.com/countries.aspx"
    HtmlText = Inet1.OpenURL(URL, icString)
    
    Set MyRegExp = New RegExp
    
    MyRegExp.Pattern = "selecttown.asp\?CountryID=(..)\"">(.*?)<\/a>"
    
    ' search for any matches withing s
    MyRegExp.Global = True
    MyRegExp.IgnoreCase = True
    MyRegExp.MultiLine = True
    
    ' return a collection of matches
    Set Matches = MyRegExp.Execute(HtmlText)
        
    If Matches.Count < 1 Then Exit Sub
    
    lvwCountries.ListItems.Clear
     
    FileNum = FreeFile
    Open AppPath + "Countries.txt" For Output As FileNum
    
    For Each Match In Matches
        CountryCode = Match.SubMatches(0)
        CountryName = Match.SubMatches(1)
        CountryName = Replace(CountryName, "&nbsp;", " ")
        Set itmX = lvwCountries.ListItems.Add(, CountryCode, CountryName)
        Print #FileNum, CountryCode & "=" & CountryName
    Next
    
    Close #FileNum
    
End Sub

Private Sub Form_Resize()
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSettings AppRegPath, "Settings", "MainLeft", Me.Left
        SaveSettings AppRegPath, "Settings", "MainTop", Me.Top
        SaveSettings AppRegPath, "Settings", "MainWidth", Me.Width
        SaveSettings AppRegPath, "Settings", "MainHeight", Me.Height
    End If

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

   ' Retrieve server response using the GetChunk
   ' method when State = 12. This example assumes the
   ' data is text.

   Select Case State
        'Case icHostResolvingHost:
            'StatusBar.Panels(1).Text = "icHostResolvingHost"
        Case icHostResolved
            StatusBar.Panels(1).Text = "Host Resolved"
        Case icConnecting
            StatusBar.Panels(1).Text = "Connecting"
        Case icConnected
            StatusBar.Panels(1).Text = "Connected"
        Case icRequesting
            StatusBar.Panels(1).Text = "Requesting"
        Case icRequestSent
            StatusBar.Panels(1).Text = "Request Sent"
        Case icReceivingResponse
            StatusBar.Panels(1).Text = "Receiving Response"
        Case icResponseReceived
            StatusBar.Panels(1).Text = "Response Received"
        Case icDisconnecting
            StatusBar.Panels(1).Text = "Disconnecting"
        Case icDisconnected
            StatusBar.Panels(1).Text = "Disconnected"
        Case icError
            StatusBar.Panels(1).Text = "Error"
        Case icResponseCompleted
            StatusBar.Panels(1).Text = "Response Completed"

   End Select

End Sub

Private Sub Label13_Click()
    txtFormat.SelText = vbTab
End Sub

Private Sub Label3_Click()
    txtFormat.SelText = "%Country%"
End Sub

Private Sub Label4_Click()
    txtFormat.SelText = "%Region%"
End Sub

Private Sub Label5_Click()
    txtFormat.SelText = "%County%"
End Sub

Private Sub Label6_Click()
    txtFormat.SelText = "%Longitude%"
End Sub

Private Sub Label7_Click()
    txtFormat.SelText = "%Latitude%"
End Sub

Private Sub Label8_Click()
    txtFormat.SelText = "%Altitude%"
End Sub

Private Sub lvwCountries_Click()
    lvwCountries.SelectedItem.Checked = Not lvwCountries.SelectedItem.Checked
End Sub

Private Sub Inet_StateChanged(Index As Integer, ByVal State As Integer)

   ' Retrieve server response using the GetChunk
   ' method when State = 12.

   Dim vtData As Variant ' Data variable.
   
   Select Case State
        ' ... Other cases not shown.
   Case icError ' 11
        ' In case of error, return ResponseCode and
        ' ResponseInfo.
        vtData = Inet1.ResponseCode & ":" & Inet1.ResponseInfo
        
   Case icResponseCompleted ' 12
        Dim bDone As Boolean
        
        bDone = False
        
        ' Get first chunk.
        'vtData = Inet(Index).GetChunk(8192, icString)
        
        DoEvents

        Do While Not bDone
            'InetHtml(Index) = InetHtml(Index) & vtData
            ' Get next chunk.
            'vtData = Inet(Index).GetChunk(1024, icString)
            DoEvents
            
            If Len(vtData) = 0 Then
               bDone = True
            End If
        Loop
        
        'InetStatus(Index) = True
   
   End Select
    
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    'StatusBar.Panels(1).Text = Text
    
End Sub

Private Sub cmdDataFolderBrowse_Click()
    Dim strDirName As String

    On Error GoTo ErrHandler
    
    strDirName = txtDataFolder.Text
    If (strDirName) = "" Then
        strDirName = AppPath + "data"
    Else
        'strDirName = Dir(strDirName, vbDirectory)
    End If
    
    'strDirName = Browse_Folder(Me.hWnd, strDirName)
    strDirName = BrowseForFolder("Select folder", strDirName, , True, False, Me)
    
    ''We don't want to have an error if the file doesn't exist.
    If strDirName <> "" Then
        strDirName = ReturnPath(strDirName)
        txtDataFolder.Text = strDirName
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Private Sub cmdSaveSettings_Click()
    OutputFormat = txtFormat.Text
    SaveSettings AppRegPath, "Settings", "OutputFormat", txtFormat.Text
    OutputFolder = ReturnPath(Trim(txtDataFolder.Text))
    txtDataFolder.Text = OutputFolder
    SaveSettings AppRegPath, "Settings", "OutputFolder", OutputFolder
End Sub

Private Sub cmdOpenDataFolder_Click()
    OpenFolder OutputFolder
End Sub

Private Sub lblMewsoft_Click()
    Dim URL As String
    URL = Join(Array("h", "t", "t", "p", ":", "/", "/", "w", "w", "w", ".", "m", "e", "w", "s", "o", "f", "t", ".", "c", "o", "m"), "~")
    URL = Replace(URL, "~", "")
    ShellDocument URL
End Sub

Private Sub lblSalesEmail_Click()
    Dim URL As String
    URL = Join(Array("m", "a", "i", "l", "t", "o", ":", "s", "a", "l", "e", "s", "@", "m", "e", "w", "s", "o", "f", "t", ".", "c", "o", "m"), "")
    URL = Replace(URL, "~", "")
    ShellDocument URL
End Sub

Private Sub lblSupportEmail_Click()
    Dim URL As String
    URL = Join(Array("m", "a", "i", "l", "t", "o", ":", "s", "u", "p", "p", "o", "r", "t", "@", "m", "e", "w", "s", "o", "f", "t", ".", "c", "o", "m"), "")
    URL = Replace(URL, "~", "")
    ShellDocument URL
End Sub


