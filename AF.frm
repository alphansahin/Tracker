VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formTracker 
   Caption         =   "Tracker"
   ClientHeight    =   12360
   ClientLeft      =   1995
   ClientTop       =   990
   ClientWidth     =   16530
   KeyPreview      =   -1  'True
   LinkTopic       =   "formTracker"
   ScaleHeight     =   12360
   ScaleWidth      =   16530
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox trackerDriveObj 
      Height          =   315
      Left            =   17040
      TabIndex        =   50
      Top             =   840
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.FileListBox trackerFileObj 
      Height          =   2625
      Left            =   19560
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.DirListBox trackerDirObj 
      Height          =   2790
      Left            =   17040
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox listAFDfilesForCabinets 
      Height          =   1815
      Left            =   16920
      TabIndex        =   46
      Top             =   9240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox listAFDFoldersForCabinets 
      Height          =   1815
      Left            =   19800
      TabIndex        =   45
      Top             =   9240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox listAFDfilesForChronicles 
      Height          =   1815
      Left            =   16920
      TabIndex        =   44
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox listAFDfoldersForChronicles 
      Height          =   1815
      Left            =   19800
      TabIndex        =   43
      Top             =   7320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox listAFDFoldersForDocuments 
      Height          =   1815
      Left            =   19800
      TabIndex        =   42
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox listAFDfilesForDocuments 
      Height          =   1815
      Left            =   16920
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame frameSearchBox 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   360
      TabIndex        =   34
      Top             =   8040
      Width           =   14475
      Begin VB.TextBox textSearchInDocuments 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   40
         Top             =   0
         Width           =   2760
      End
      Begin VB.TextBox textSearchInChronicles 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   39
         Top             =   0
         Width           =   3000
      End
      Begin VB.TextBox textSearchInCabinets 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11400
         TabIndex        =   38
         Top             =   0
         Width           =   3000
      End
      Begin VB.Label labelSearchInChronicles 
         Caption         =   "Search in Chronicles:"
         Height          =   210
         Left            =   4800
         TabIndex        =   37
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label labelSearchInCabinets 
         Caption         =   "Search in Cabinets:"
         Height          =   210
         Left            =   9840
         TabIndex        =   36
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label labelSearchInDocuments 
         Caption         =   "Search in Documents:"
         Height          =   210
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSComctlLib.ListView listViewCabinets 
      Height          =   5295
      Left            =   9600
      TabIndex        =   33
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9340
      View            =   3
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Score"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Title"
         Object.Width           =   13230
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Author"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Venue"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Extension"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Path"
         Object.Width           =   4288
      EndProperty
   End
   Begin MSComctlLib.ListView listViewChronicles 
      Height          =   5295
      Left            =   4440
      TabIndex        =   32
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9340
      View            =   3
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Score"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Title"
         Object.Width           =   13230
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Author"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Venue"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Extension"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Path"
         Object.Width           =   4288
      EndProperty
   End
   Begin MSComctlLib.ListView ListViewDocuments 
      Height          =   5295
      Left            =   360
      TabIndex        =   31
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9340
      View            =   3
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Score"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Title"
         Object.Width           =   11466
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Author"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Venue"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Extension"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Path"
         Object.Width           =   4288
      EndProperty
   End
   Begin MSComctlLib.ProgressBar progressBarUpdateList 
      Height          =   165
      Left            =   14760
      TabIndex        =   29
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdActiveFolders 
      Caption         =   "Active Folders"
      Height          =   375
      Left            =   14760
      TabIndex        =   30
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdateList 
      Caption         =   "Update List"
      Height          =   492
      Left            =   14760
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame frameDetails 
      BorderStyle     =   0  'None
      Height          =   3012
      Left            =   240
      TabIndex        =   10
      Top             =   8880
      Width           =   16095
      Begin VB.TextBox textScore 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13560
         TabIndex        =   11
         Top             =   2160
         Width           =   2532
      End
      Begin VB.CommandButton cmdUpdateDetails 
         Caption         =   "Update Details"
         Height          =   495
         Left            =   12360
         TabIndex        =   20
         Top             =   2520
         Width           =   3735
      End
      Begin VB.OptionButton opt_abstract 
         Caption         =   "Abstract"
         Height          =   255
         Left            =   13560
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton opt_unreviewed 
         Caption         =   "Unreviewed"
         Height          =   255
         Left            =   13560
         TabIndex        =   18
         Top             =   1680
         Width           =   1212
      End
      Begin VB.OptionButton opt_skim 
         Caption         =   "Skim"
         Height          =   255
         Left            =   14760
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox textVenue 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   13560
         TabIndex        =   16
         Top             =   1320
         Width           =   2532
      End
      Begin VB.TextBox textYear 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   13560
         TabIndex        =   15
         Top             =   960
         Width           =   2532
      End
      Begin VB.TextBox textAuthor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   13560
         TabIndex        =   14
         Top             =   600
         Width           =   2532
      End
      Begin VB.TextBox textTitle 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   492
         Left            =   13560
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   0
         Width           =   2532
      End
      Begin VB.OptionButton opt_read 
         Caption         =   "Read"
         Height          =   255
         Left            =   14760
         TabIndex        =   12
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox textDetails 
         Appearance      =   0  'Flat
         Height          =   3015
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   0
         Width           =   12255
      End
      Begin VB.Label labelDetailsScore 
         Caption         =   "Score (0-100):"
         Height          =   255
         Left            =   12360
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label labelDetailsVenue 
         Caption         =   "Venue:"
         Height          =   255
         Left            =   12360
         TabIndex        =   25
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label labelDetailsYear 
         Caption         =   "Year:"
         Height          =   255
         Left            =   12360
         TabIndex        =   24
         Top             =   960
         Width           =   495
      End
      Begin VB.Label labelDetailsAuthor 
         Caption         =   "Author:"
         Height          =   255
         Left            =   12360
         TabIndex        =   23
         Top             =   600
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label labelDetailsTitle 
         Caption         =   "Title:"
         Height          =   255
         Left            =   12360
         TabIndex        =   22
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label labelDetailsReview 
         Caption         =   "Review:"
         Height          =   255
         Left            =   12360
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdNewChronicle 
      Caption         =   "New Chronicle"
      Height          =   375
      Left            =   14760
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewCabinet 
      Caption         =   "New Cabinet"
      Height          =   375
      Left            =   14760
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdFileExport 
      Caption         =   "Export Checked File(s)"
      Height          =   555
      Left            =   14760
      TabIndex        =   5
      Top             =   3960
      Width           =   1545
   End
   Begin VB.CommandButton cmdFileDelete 
      Caption         =   "Delete Checked File(s)"
      Height          =   495
      Left            =   14760
      TabIndex        =   2
      Top             =   4560
      Width           =   1545
   End
   Begin VB.CommandButton cmdNewNote 
      Caption         =   "New Note"
      Height          =   375
      Left            =   14760
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewDocument 
      Caption         =   "New Document"
      Height          =   375
      Left            =   14760
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewRecording 
      Caption         =   "New Recording"
      Height          =   375
      Left            =   14760
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip tabDetails 
      Height          =   3495
      Left            =   100
      TabIndex        =   28
      Top             =   8500
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   6165
      TabMinWidth     =   706
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Details"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   8295
      Left            =   100
      TabIndex        =   6
      Top             =   120
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   14631
      TabWidthStyle   =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Documents"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Chronicles"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cabinets"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labelStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   100
      TabIndex        =   47
      Top             =   12100
      Width           =   11295
   End
   Begin VB.Label labelCurrentDir 
      Height          =   255
      Left            =   11520
      TabIndex        =   0
      Top             =   12100
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "formTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©2010-2018 Alphan Sahin, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   FormControl Version 2.0
'   Code module for resizing a form based on screen size, then resizing the
'   controls based on the forms size
'
'   Copyright (C) 2007
'   Richard L. McCutchen
'   Email: richard@psychocoder.net
'   Created: AUG99
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.

Private Const vbDot = 46
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const vbBackslash = "\"
Private Const ALL_FILES = "*.*"

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

Private Type FILE_PARAMS
   bRecurse As Boolean
   nCount As Long
   nSearched As Long
   sFileNameExt As String
   sFileRoot As String
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function lstrlen Lib "kernel32" _
    Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Declare Function PathMatchSpec Lib "shlwapi" _
   Alias "PathMatchSpecW" _
  (ByVal pszFileParam As Long, _
   ByVal pszSpec As Long) As Long

Private fp As FILE_PARAMS  'holds search parameters

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_NEWDIALOGSTYLE As Long = &H40



Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Public strdoc, strcho, strsof, varTextDetailsChange
Public savepath, saveart_name, saveaut_name, saveart_date, savefile_type, savedate_format
Public another
Public IsAFDpathsForDocumentsUpdated, IsAFDpathsForChroniclesUpdated, IsAFDpathsForCabinetsUpdated

Private list() As Controlx

Private Type Controlx
    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
    iWidth As Integer
    iHeight As Integer
End Type

Private Declare Function SetCurrentDirectory Lib "kernel32" _
Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long





Public Function DirName(loc As String)

For i = 1 To Len(loc)
    dot_point = InStr(Len(loc) - i, loc, "\")
    If dot_point = 0 Then
    
    Else
        DirName = Mid$(loc, dot_point + 1, Len(loc))
        Exit Function
    End If
Next i


End Function

Public Function TypeName(loc As String)
dot_point = InStr(1, StrReverse(loc), ".")
TypeName = Mid$(loc, Len(loc) - (dot_point - 2), Len(loc))

End Function


Public Function makedir(loc As String)
On Local Error GoTo Hata

MkDir (loc)
Hata:
End Function


Private Sub cmdNewDocument_Click()
Load formAddDocument
formAddDocument.Show
End Sub


Private Sub cmdNewCabinet_Click()
Load formAddCabinet
formAddCabinet.Show
End Sub

Private Sub cmdNewChronicle_Click()
Load formAddChronicles
formAddChronicles.Show
End Sub


Private Sub cmdActiveFolders_Click()
Load formActiveFolders
formActiveFolders.Show
End Sub


Private Sub cmdUpdateList_Click()
cmdUpdateList.Caption = "Wait..."
cmdUpdateList.Enabled = 0
cmdFileDelete.Enabled = 0
cmdFileExport.Enabled = 0

If (tabMain.SelectedItem.Index) = 1 Then
    Call new_ordertree(True, False, False)
ElseIf (tabMain.SelectedItem.Index) = 2 Then
    Call new_ordertree(False, True, False)
ElseIf (tabMain.SelectedItem.Index) = 3 Then
    Call new_ordertree(False, False, True)
End If
    
cmdUpdateList.Enabled = 1
cmdFileDelete.Enabled = 1
cmdFileExport.Enabled = 1
cmdUpdateList.Caption = "Update List"
End Sub

Private Sub cmdFileDelete_Click()

cmdFileExport.Enabled = 0
cmdFileDelete.Enabled = 0
cmdUpdateList.Enabled = 0
cmdFileDelete.Caption = "Wait..."


Dim docc As Boolean
Dim chro As Boolean
Dim soft As Boolean

docc = False
chro = False
soft = False

labelStatus = "Status: Deleting document(s)..."
DoEvents
For i = ListViewDocuments.ListItems.Count To 1 Step -1
    Set ListViewDocuments.SelectedItem = ListViewDocuments.ListItems(i)
    If ListViewDocuments.SelectedItem.Checked = True Then
        docc = True
        IsAFDpathsForDocumentsUpdated = 0
    
        date_format = ListViewDocuments.SelectedItem
        art_name = ListViewDocuments.SelectedItem.ListSubItems(4)
        aut_name = ListViewDocuments.SelectedItem.ListSubItems(5)
        art_date = ListViewDocuments.SelectedItem.ListSubItems(6)
        file_type = ListViewDocuments.SelectedItem.ListSubItems(8)
    
        trackerDirObj.Path = ListViewDocuments.SelectedItem.ListSubItems(9)
        trackerFileObj.Path = ListViewDocuments.SelectedItem.ListSubItems(9)
        If trackerFileObj.ListCount = 2 Then
            Kill ListViewDocuments.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
            Kill ListViewDocuments.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            RmDir removeLastCharacter(ListViewDocuments.SelectedItem.ListSubItems(9))
            trackerDirObj.Path = labelCurrentDir & "\documents" & "\" & date_format
            If trackerDirObj.ListCount = 0 Then
                trackerDirObj.Path = labelCurrentDir & "\documents"
                RmDir labelCurrentDir & "\documents" & "\" & date_format
                trackerDirObj.Refresh
            End If
        Else
            Kill ListViewDocuments.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
            Kill ListViewDocuments.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            trackerDirObj.Path = labelCurrentDir & "\documents\" & date_format
            If trackerDirObj.ListCount = 0 Then
                trackerDirObj.Path = labelCurrentDir & "\documents"
                RmDir labelCurrentDir & "\documents" & "\" & date_format
                trackerDirObj.Refresh
            End If
        End If
        ListViewDocuments.ListItems.Remove ListViewDocuments.SelectedItem.Index
        formTracker.IsAFDpathsForDocumentsUpdated = 0
    End If
Next i
tabMain.Tabs(1).Caption = "Documents (" & ListViewDocuments.ListItems.Count & ")"


For i = listViewChronicles.ListItems.Count To 1 Step -1
    Set listViewChronicles.SelectedItem = listViewChronicles.ListItems(i)
    If listViewChronicles.SelectedItem.Checked = True Then
        chro = True
        IsAFDpathsForChroniclesUpdated = 0
        
        date_format = listViewChronicles.SelectedItem
        art_name = listViewChronicles.SelectedItem.ListSubItems(4)
        aut_name = listViewChronicles.SelectedItem.ListSubItems(5)
        art_date = listViewChronicles.SelectedItem.ListSubItems(6)
        file_type = listViewChronicles.SelectedItem.ListSubItems(8)
    
        trackerDirObj.Path = listViewChronicles.SelectedItem.ListSubItems(9)
        trackerFileObj.Path = listViewChronicles.SelectedItem.ListSubItems(9)
        If trackerFileObj.ListCount = 2 Then
            Kill listViewChronicles.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
            Kill listViewChronicles.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            RmDir removeLastCharacter(listViewChronicles.SelectedItem.ListSubItems(9))
            trackerDirObj.Path = labelCurrentDir & "\chronicles" & "\" & date_format
            If trackerDirObj.ListCount = 0 Then
                trackerDirObj.Path = labelCurrentDir & "\chronicles"
                RmDir labelCurrentDir & "\chronicles" & "\" & date_format
                trackerDirObj.Refresh
            End If
        Else
            Kill listViewChronicles.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
            Kill listViewChronicles.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            trackerDirObj.Path = labelCurrentDir & "\chronicles\" & date_format
            If trackerDirObj.ListCount = 0 Then
                trackerDirObj.Path = labelCurrentDir & "\chronicles"
                RmDir labelCurrentDir & "\chronicles" & "\" & date_format
                trackerDirObj.Refresh
            End If
        End If
        listViewChronicles.ListItems.Remove listViewChronicles.SelectedItem.Index
        formTracker.IsAFDpathsForChroniclesUpdated = 0
    End If
Next i
tabMain.Tabs(2).Caption = "Chronicles (" & listViewChronicles.ListItems.Count & ")"

For i = listViewCabinets.ListItems.Count To 1 Step -1
    Set listViewCabinets.SelectedItem = listViewCabinets.ListItems(i)
    If listViewCabinets.SelectedItem.Checked = True Then
    
        soft = True
        
        IsAFDpathsForCabinetsUpdated = 0
        
        date_format = listViewCabinets.SelectedItem
        art_name = listViewCabinets.SelectedItem.ListSubItems(4)
        aut_name = listViewCabinets.SelectedItem.ListSubItems(5)
        art_date = listViewCabinets.SelectedItem.ListSubItems(6)
        file_type = listViewCabinets.SelectedItem.ListSubItems(8)
    
        trackerDirObj.Path = listViewCabinets.SelectedItem.ListSubItems(9)
        trackerFileObj.Path = listViewCabinets.SelectedItem.ListSubItems(9)
        If trackerFileObj.ListCount = 2 Then
            Kill listViewCabinets.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
            Kill listViewCabinets.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            RmDir removeLastCharacter(listViewCabinets.SelectedItem.ListSubItems(9))
            trackerDirObj.Path = labelCurrentDir & "\cabinets" & "\" & date_format
            If trackerDirObj.ListCount = 0 Then
                trackerDirObj.Path = labelCurrentDir & "\cabinets"
                RmDir labelCurrentDir & "\cabinets" & "\" & date_format
                trackerDirObj.Refresh
            End If
        Else
            Kill listViewCabinets.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
            Kill listViewCabinets.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            trackerDirObj.Path = labelCurrentDir & "\cabinets" & "\" & date_format
            If trackerDirObj.ListCount = 0 Then
                trackerDirObj.Path = labelCurrentDir & "\cabinets   "
                RmDir labelCurrentDir & "\cabinets" & "\" & date_format
                trackerDirObj.Refresh
            End If
        End If
        listViewCabinets.ListItems.Remove listViewCabinets.SelectedItem.Index
        formTracker.IsAFDpathsForCabinetsUpdated = 0
    End If
Next i
tabMain.Tabs(3).Caption = "Cabinets (" & listViewCabinets.ListItems.Count & ")"

labelStatus = "Status: File(s) are deleted."
DoEvents


cmdFileDelete.Caption = "Delete"
cmdFileDelete.Enabled = 1
cmdUpdateList.Enabled = 1
cmdFileExport.Enabled = 1

End Sub



Private Sub cmdUpdateDetails_Click()
Call savecomments
varTextDetailsChange = textDetails.Text
End Sub


Private Sub cmdNewNote_Click()
Load formAddNote
formAddNote.Show
End Sub

Private Sub cmdNewRecording_Click()
Load formAddRecoding
formAddRecoding.Show
End Sub

Private Sub cmdFileExport_Click()
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

szTitle = "Select directory for extracting..."
With tBrowseInfo
   .hwndOwner = Me.hwnd
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_NEWDIALOGSTYLE
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   
   

    For i = 1 To ListViewDocuments.ListItems.Count
        Set ListViewDocuments.SelectedItem = ListViewDocuments.ListItems(i)
        If ListViewDocuments.SelectedItem.Checked = True Then
            date_format = ListViewDocuments.SelectedItem
            art_name = ListViewDocuments.SelectedItem.ListSubItems(4)
            aut_name = ListViewDocuments.SelectedItem.ListSubItems(5)
            art_date = ListViewDocuments.SelectedItem.ListSubItems(6)
            file_type = ListViewDocuments.SelectedItem.ListSubItems(8)
        

            SourceFile = ListViewDocuments.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            DestinationFile = sBuffer & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            FileCopy SourceFile, DestinationFile
        End If
    Next i
    
    For i = 1 To listViewChronicles.ListItems.Count
        Set listViewChronicles.SelectedItem = listViewChronicles.ListItems(i)
        If listViewChronicles.SelectedItem.Checked = True Then
            date_format = listViewChronicles.SelectedItem
            art_name = listViewChronicles.SelectedItem.ListSubItems(4)
            aut_name = listViewChronicles.SelectedItem.ListSubItems(5)
            art_date = listViewChronicles.SelectedItem.ListSubItems(6)
            file_type = listViewChronicles.SelectedItem.ListSubItems(8)
        

            SourceFile = listViewChronicles.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            DestinationFile = sBuffer & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            FileCopy SourceFile, DestinationFile
        End If
    Next i
    
    For i = 1 To listViewCabinets.ListItems.Count
        Set listViewCabinets.SelectedItem = listViewCabinets.ListItems(i)
        If listViewCabinets.SelectedItem.Checked = True Then
            date_format = listViewCabinets.SelectedItem
            art_name = listViewCabinets.SelectedItem.ListSubItems(4)
            aut_name = listViewCabinets.SelectedItem.ListSubItems(5)
            art_date = listViewCabinets.SelectedItem.ListSubItems(6)
            file_type = listViewCabinets.SelectedItem.ListSubItems(8)
        
            
            SourceFile = listViewCabinets.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            DestinationFile = sBuffer & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            FileCopy SourceFile, DestinationFile
        End If
    Next i
    
   
End If
MsgBox "File(s) are exported. Path:" & sBuffer

End Sub


Private Sub Form_Load()
strdoc = ""
strcho = ""
strsof = ""

'SetCurrentDirectory "C:\Dropbox\Work\Project_Tracker\code"

labelCurrentDir = CurDir

txtFolder = "documents"
Call obtainAFDList(txtFolder, listAFDfilesForDocuments, listAFDFoldersForDocuments)
txtFolder = "chronicles"
Call obtainAFDList(txtFolder, listAFDfilesForChronicles, listAFDfoldersForChronicles)
txtFolder = "cabinets"
Call obtainAFDList(txtFolder, listAFDfilesForCabinets, listAFDFoldersForCabinets)
IsAFDpathsForDocumentsUpdated = 1
IsAFDpathsForChroniclesUpdated = 1
IsAFDpathsForCabinetsUpdated = 1


valLeft = 240
valTop = 500
valWidth = 14415
valHeight = 7435
ListViewDocuments.Top = valTop
ListViewDocuments.width = valWidth
ListViewDocuments.Left = valLeft
ListViewDocuments.height = valHeight

listViewChronicles.Top = valTop
listViewChronicles.width = valWidth
listViewChronicles.Left = valLeft
listViewChronicles.height = valHeight

listViewCabinets.Top = valTop
listViewCabinets.width = valWidth
listViewCabinets.Left = valLeft
listViewCabinets.height = valHeight

GetLocation formTracker


Call new_ordertree(True, True, True)
formTracker.Show

Call doc_vis

End Sub




Public Function webtree_sub(list As ListView, searchArray)
Dim FSO As New FileSystemObject
Set fld = FSO.GetFolder(Dir1)
list.ListItems.Clear
list.Visible = False
'On Error GoTo Hata

ProgressBar1.Min = 0
ProgressBar1.Value = 0
For Each sfld In fld.SubFolders
    ProgressBar1.Max = fld.SubFolders.Count
    ProgressBar1.Value = ProgressBar1.Value + 1
    For Each ssfld In sfld.SubFolders
         For Each fl In ssfld.Files
            If TypeName(fl.Name) = "afd" Then

                Open fl.Path For Input As 3
                    Input #3, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
                Close #3
                
                add_ListOR = 0
                For indSearch = 0 To UBound(searchArray)
                    sttrAll = searchArray(indSearch)
                    add_listAND = 1
                    
                    startPos = 1
                    Do While startPos < Len(sttrAll)
                        stopPos = InStr(startPos, sttrAll, ",")
                        strr = Mid(sttrAll, startPos, stopPos - startPos)
                        startPos = stopPos + 1
                    
                            
                        'If InStr(1, UCase(from_name), UCase(strr)) > 0 Then
                        If InStrB(UCase(from_name), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(date_format), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(date_format), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(art_name), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(art_name), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(from_name), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(from_name), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(aut_name), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(aut_name), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(art_conf), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(art_conf), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(art_date), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(art_date), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(art_keys), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(art_keys), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(file_type), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(file_type), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(review_score), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(review_score), UCase(strr)) <> 0 Then
                            add_listR = 1
                        'ElseIf InStr(1, UCase(review_status), UCase(strr)) > 0 Then
                        ElseIf InStrB(UCase(review_status), UCase(strr)) <> 0 Then
                            add_listR = 1
                            
                        Else
                            add_listR = 0
                        End If
                        
                        add_listAND = add_listAND And add_listR
                    Loop
                    add_ListOR = add_ListOR Or add_listAND
                Next
                add_list = add_ListOR
                
                If add_list = 1 Then
                    Set lv = list.ListItems.Add(1, , date_format, , 0)
                    lv.ListSubItems.Add , , from_name
                    If review_status = "1" Then
                        re_status = "Unreviewed"
                    ElseIf review_status = "2" Then
                        re_status = "Abstract"
                    ElseIf review_status = "3" Then
                        re_status = "Skim"
                    ElseIf review_status = "4" Then
                        re_status = "Read"
                    End If
                    lv.ListSubItems.Add , , re_status
                    lv.ListSubItems.Add , , review_score
                    lv.ListSubItems.Add , , art_name
                    lv.ListSubItems.Add , , aut_name
                    lv.ListSubItems.Add , , art_date
                    lv.ListSubItems.Add , , art_conf
                    lv.ListSubItems.Add , , LCase(file_type)
                    lv.ListSubItems.Add , , fl.ParentFolder.Path
                End If
                DoEvents
            End If
        Next
    Next
Next
ChDrive Left$(CurrentDir, 1)
ChDir CurrentDir
list.Visible = True
Exit Function
Hata:
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
    If (tabMain.SelectedItem.Index) = 1 Then
        Call new_ordertree(True, False, False)
    ElseIf (tabMain.SelectedItem.Index) = 2 Then
        Call new_ordertree(False, True, False)
    ElseIf (tabMain.SelectedItem.Index) = 3 Then
        Call new_ordertree(False, False, True)
    End If
End If
End Sub


Private Sub trackerDirObj_Change()
trackerFileObj.Path = trackerDirObj.Path

End Sub

Private Sub trackerDriveObj_Change()
trackerDirObj.Path = trackerDriveObj.Drive

End Sub


Public Function doc_vis()
    ListViewDocuments.Visible = True
    listViewChronicles.Visible = False
    listViewCabinets.Visible = False
End Function

Public Function no_vis()
    ListViewDocuments.Visible = False
    listViewChronicles.Visible = False
    listViewCabinets.Visible = False
End Function

Public Function cro_vis()
    ListViewDocuments.Visible = False
    listViewChronicles.Visible = True
    listViewCabinets.Visible = False
End Function

Public Function sof_vis()
    ListViewDocuments.Visible = False
    listViewChronicles.Visible = False
    listViewCabinets.Visible = True
End Function



Public Function savecomments()


If opt_unreviewed = True Then
    yreview_status = "1"
ElseIf opt_abstract = True Then
    yreview_status = "2"
ElseIf opt_skim = True Then
    yreview_status = "3"
ElseIf opt_read = True Then
    yreview_status = "4"
End If

Open savepath & saveart_name & "_" & saveaut_name & "_" & saveart_date & "_" & savefile_type & ".afd" For Input As 1
Open savepath & "temp" For Output As 2
    i = 1
    Do Until EOF(1)
    Input #1, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
    If i = 1 Then
        Write #2, from_name, date_format, art_name, textAuthor, textVenue, textYear, textDetails.Text, file_type, textScore.Text, yreview_status, directoryy
    Else
        Write #2, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
    End If
    i = i + 1
    Loop
Close

Kill savepath & saveart_name & "_" & saveaut_name & "_" & saveart_date & "_" & savefile_type & ".afd"
Name savepath & "temp" As savepath & saveart_name & "_" & textAuthor & "_" & textYear & "_" & savefile_type & ".afd"
End Function


Public Function obtainAFDList(txtFolder, listForFiles As ListBox, listForFolders As ListBox)
    listForFiles.Clear
    listForFolders.Clear
    With fp
      .sFileRoot = labelCurrentDir & "\" & txtFolder & "\" 'start path
      .sFileNameExt = "*.afd"           'file type(s) of interest
      .bRecurse = True         'True = recursive search
      .nCount = 0                          'results
      .nSearched = 0                       'results
   End With
  
   Call SearchForFiles(fp.sFileRoot, listForFiles, listForFolders)
   
   nSearched = Format$(fp.nSearched, "###,###,###,##0")
   nCount = Format$(fp.nCount, "###,###,###,##0")

End Function





Public Function activeFolderSearch(searchArray)
cmdFileExport.Enabled = 0
cmdFileDelete.Enabled = 0
cmdUpdateList.Enabled = 0

stateDocumentsVisible = ListViewDocuments.Visible
stateChroniclesVisible = listViewChronicles.Visible
stateCabinetsVisible = listViewCabinets.Visible
Call no_vis

ListViewDocuments.Enabled = False
listViewChronicles.Enabled = False
listViewCabinets.Enabled = False

labelStatus = "Status: Updating the list for documents..."
strFolder = "documents"

progressBarUpdateList.Value = 0
progressBarUpdateList.Visible = True
ListViewDocuments.ListItems.Clear

If IsAFDpathsForDocumentsUpdated = 0 Then
    labelStatusPre = labelStatus
    labelStatus = "Status: Refreshing AFD list..."
    DoEvents
    Call obtainAFDList(strFolder, listAFDfilesForDocuments, listAFDFoldersForDocuments)
    IsAFDpathsForDocumentsUpdated = 1
    labelStatus = labelStatusPre
    DoEvents
End If
Call new_webtree_sub(ListViewDocuments, searchArray, listAFDfilesForDocuments, listAFDFoldersForDocuments)
progressBarUpdateList.Visible = False

tabMain.Tabs(1).Caption = "Documents (" & ListViewDocuments.ListItems.Count & ")"
    
cmdUpdateList.Enabled = 1
cmdUpdateList.Caption = "Update List"


ListViewDocuments.Enabled = True
listViewChronicles.Enabled = True
listViewCabinets.Enabled = True

ListViewDocuments.Visible = stateDocumentsVisible
listViewChronicles.Visible = stateChroniclesVisible
listViewCabinets.Visible = stateCabinetsVisible

cmdFileExport.Enabled = 1
cmdFileDelete.Enabled = 1
cmdUpdateList.Enabled = 1

labelStatus = "Status: The list(s) are updated"
End Function




Public Function new_ordertree(forceSearchDocuments As Boolean, forceSearchChronicles As Boolean, forceSearchCabinets As Boolean)



stateDocumentsVisible = ListViewDocuments.Visible
stateChroniclesVisible = listViewChronicles.Visible
stateCabinetsVisible = listViewCabinets.Visible
Call no_vis

ListViewDocuments.Enabled = False
listViewChronicles.Enabled = False
listViewCabinets.Enabled = False

If forceSearchDocuments = True Then
    labelStatus = "Status: Updating the list for documents..."
    strFolder = "documents"
    Call doSearchList(ListViewDocuments, textSearchInDocuments.Text, strFolder, IsAFDpathsForDocumentsUpdated, listAFDfilesForDocuments, listAFDFoldersForDocuments)
    tabMain.Tabs(1).Caption = "Documents (" & ListViewDocuments.ListItems.Count & ")"
 Else
    If strdoc <> textSearchInDocuments.Text Then
        labelStatus = "Status: Updating the list for documents..."
        strFolder = "documents"
        Call doSearchList(ListViewDocuments, textSearchInDocuments.Text, strFolder, IsAFDpathsForDocumentsUpdated, listAFDfilesForDocuments, listAFDFoldersForDocuments)
        tabMain.Tabs(1).Caption = "Documents (" & ListViewDocuments.ListItems.Count & ")"
    End If
End If
    
If forceSearchChronicles = True Then
    labelStatus = "Status: Updating the list for chronicles..."
    strFolder = "chronicles"
    Call doSearchList(listViewChronicles, textSearchInChronicles.Text, strFolder, IsAFDpathsForChroniclesUpdated, listAFDfilesForChronicles, listAFDfoldersForChronicles)
    tabMain.Tabs(2).Caption = "Chronicles (" & listViewChronicles.ListItems.Count & ")"
Else
    If strcho <> textSearchInChronicles.Text Then
        labelStatus = "Status: Updating the list for chronicles..."
        strFolder = "chronicles"
        Call doSearchList(listViewChronicles, textSearchInChronicles.Text, strFolder, IsAFDpathsForChroniclesUpdated, listAFDfilesForChronicles, listAFDfoldersForChronicles)
        tabMain.Tabs(2).Caption = "Chronicles (" & listViewChronicles.ListItems.Count & ")"
    End If
End If

If forceSearchCabinets = True Then
    labelStatus = "Status: Updating the list for cabinets..."
    strFolder = "cabinets"
    Call doSearchList(listViewCabinets, textSearchInCabinets.Text, strFolder, IsAFDpathsForCabinetsUpdated, listAFDfilesForCabinets, listAFDFoldersForCabinets)
    tabMain.Tabs(3).Caption = "Cabinets (" & listViewCabinets.ListItems.Count & ")"
Else
    If strsof <> textSearchInCabinets.Text Then
        labelStatus = "Status: Updating the list for cabinets..."
        strFolder = "cabinets"
        Call doSearchList(listViewCabinets, textSearchInCabinets.Text, strFolder, IsAFDpathsForCabinetsUpdated, listAFDfilesForCabinets, listAFDFoldersForCabinets)
        tabMain.Tabs(3).Caption = "Cabinets (" & listViewCabinets.ListItems.Count & ")"
    End If
End If

strdoc = textSearchInDocuments.Text
strcho = textSearchInChronicles.Text
strsof = textSearchInCabinets.Text




ListViewDocuments.Enabled = True
listViewChronicles.Enabled = True
listViewCabinets.Enabled = True

ListViewDocuments.Visible = stateDocumentsVisible
listViewChronicles.Visible = stateChroniclesVisible
listViewCabinets.Visible = stateCabinetsVisible

labelStatus = "Status: The list(s) are updated"
End Function

Public Function doSearchList(listForInterface As ListView, keyword As String, strFolder, isAFDUpdated As Variant, listAFDfiles As ListBox, listAFDFolders As ListBox)
    progressBarUpdateList.Value = 0
    progressBarUpdateList.Visible = True
    listForInterface.ListItems.Clear

    If Dir(labelCurrentDir & "\" & strFolder, vbDirectory) = "" Then
        'directory doesn't exist...create it
        MkDir labelCurrentDir & "\" & strFolder
    End If
    
    
    If isAFDUpdated = 0 Then
        labelStatusPre = labelStatus
        labelStatus = "Status: Refreshing AFD list..."
        DoEvents
        Call obtainAFDList(strFolder, listAFDfiles, listAFDFolders)
        isAFDUpdated = 1
        labelStatus = labelStatusPre
        DoEvents
    End If
    Call new_ordertree_sub(listForInterface, keyword, listAFDfiles, listAFDFolders)
    progressBarUpdateList.Visible = False
End Function


Public Function new_ordertree_sub(list As ListView, strr As String, listAFDfiles As ListBox, listAFDFolders As ListBox)
list.ListItems.Clear

If listAFDfiles.ListCount > 0 Then

    progressBarUpdateList.Min = 0
    progressBarUpdateList.Value = 0
    progressBarUpdateList.Max = listAFDfiles.ListCount
    
    For indexAFDfile = 0 To listAFDfiles.ListCount - 1
        progressBarUpdateList.Value = progressBarUpdateList.Value + 1
        Open listAFDfiles.list(indexAFDfile) For Input As 3
            Input #3, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
        Close
        'If InStr(1, UCase(from_name), UCase(strr)) > 0 Then
        If InStrB(UCase(from_name), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(date_format), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(date_format), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(art_name), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(art_name), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(from_name), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(from_name), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(aut_name), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(aut_name), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(art_conf), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(art_conf), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(art_date), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(art_date), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(art_keys), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(art_keys), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(file_type), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(file_type), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(review_score), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(review_score), UCase(strr)) <> 0 Then
            add_list = 1
        'ElseIf InStr(1, UCase(review_status), UCase(strr)) > 0 Then
        ElseIf InStrB(UCase(review_status), UCase(strr)) <> 0 Then
            add_list = 1
        Else
            add_list = 0
        End If
        
        If add_list = 1 Then
            Set lv = list.ListItems.Add(1, , date_format, , 0)
            lv.ListSubItems.Add , , from_name
            If review_status = "1" Then
                re_status = "Unreviewed"
            ElseIf review_status = "2" Then
                re_status = "Abstract"
            ElseIf review_status = "3" Then
                re_status = "Skim"
            ElseIf review_status = "4" Then
                re_status = "Read"
            End If
            lv.ListSubItems.Add , , re_status
            lv.ListSubItems.Add , , review_score
            lv.ListSubItems.Add , , art_name
            lv.ListSubItems.Add , , aut_name
            lv.ListSubItems.Add , , art_date
            lv.ListSubItems.Add , , art_conf
            lv.ListSubItems.Add , , LCase(file_type)
            lv.ListSubItems.Add , , listAFDFolders.list(indexAFDfile)
        End If
        DoEvents
    Next
End If



Exit Function
Hata:
End Function

Public Function new_webtree_sub(list As ListView, searchArray, listAFDfiles As ListBox, listAFDFolders As ListBox)
list.ListItems.Clear

progressBarUpdateList.Min = 0
progressBarUpdateList.Value = 0
progressBarUpdateList.Max = listAFDfiles.ListCount

For indexAFDfile = 0 To listAFDfiles.ListCount - 1
    If progressBarUpdateList.Value < listAFDfiles.ListCount Then
        progressBarUpdateList.Value = progressBarUpdateList.Value + 1
    End If
    Open listAFDfiles.list(indexAFDfile) For Input As 3
        Input #3, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
    Close #3
    
    add_ListOR = 0
    For indSearch = 0 To UBound(searchArray)
        sttrAll = searchArray(indSearch)
        add_listAND = 1
        
        startPos = 1
        Do While startPos < Len(sttrAll)
            stopPos = InStr(startPos, sttrAll, ",")
            strr = Mid(sttrAll, startPos, stopPos - startPos)
            startPos = stopPos + 1
        
                
            'If InStr(1, UCase(from_name), UCase(strr)) > 0 Then
            If InStrB(UCase(from_name), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(date_format), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(date_format), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(art_name), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(art_name), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(from_name), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(from_name), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(aut_name), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(aut_name), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(art_conf), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(art_conf), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(art_date), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(art_date), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(art_keys), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(art_keys), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(file_type), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(file_type), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(review_score), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(review_score), UCase(strr)) <> 0 Then
                add_listR = 1
            'ElseIf InStr(1, UCase(review_status), UCase(strr)) > 0 Then
            ElseIf InStrB(UCase(review_status), UCase(strr)) <> 0 Then
                add_listR = 1
            Else
                add_listR = 0
            End If
            
            add_listAND = add_listAND And add_listR
        Loop
        add_ListOR = add_ListOR Or add_listAND
    Next
    add_list = add_ListOR
    
    If add_list = 1 Then
        Set lv = list.ListItems.Add(1, , date_format, , 0)
        lv.ListSubItems.Add , , from_name
        If review_status = "1" Then
            re_status = "Unreviewed"
        ElseIf review_status = "2" Then
            re_status = "Abstract"
        ElseIf review_status = "3" Then
            re_status = "Skim"
        ElseIf review_status = "4" Then
            re_status = "Read"
        End If
        lv.ListSubItems.Add , , re_status
        lv.ListSubItems.Add , , review_score
        lv.ListSubItems.Add , , art_name
        lv.ListSubItems.Add , , aut_name
        lv.ListSubItems.Add , , art_date
        lv.ListSubItems.Add , , art_conf
        lv.ListSubItems.Add , , LCase(file_type)
        lv.ListSubItems.Add , , listAFDFolders.list(indexAFDfile)
    End If
    DoEvents
Next



Exit Function
Hata:
End Function


Private Sub Form_Resize()
    ResizeControls formTracker
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload formAddDocument
Unload formAddNote
Unload formAddRecoding
Unload formAddCabinet
Unload formAddChronicles
Unload formActiveFolders
End Sub


Private Sub listViewDocuments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListViewDocuments.Sorted = True
ListViewDocuments.SortKey = ColumnHeader.Index - 1
If ListViewDocuments.SortOrder = lvwDescending Then
    ListViewDocuments.SortOrder = lvwAscending
Else
    ListViewDocuments.SortOrder = lvwDescending
End If
End Sub

Private Sub listViewDocuments_DblClick()
art_name = ListViewDocuments.SelectedItem.ListSubItems(4)
aut_name = ListViewDocuments.SelectedItem.ListSubItems(5)
art_date = ListViewDocuments.SelectedItem.ListSubItems(6)
file_type = ListViewDocuments.SelectedItem.ListSubItems(8)

ShellExecute hwnd, "open", ListViewDocuments.SelectedItem.ListSubItems(9) & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(ListViewDocuments.SelectedItem.ListSubItems(9), 3), 1
End Sub


Private Sub listViewChronicles_DblClick()
art_name = listViewChronicles.SelectedItem.ListSubItems(4)
aut_name = listViewChronicles.SelectedItem.ListSubItems(5)
art_date = listViewChronicles.SelectedItem.ListSubItems(6)
file_type = listViewChronicles.SelectedItem.ListSubItems(8)

ShellExecute hwnd, "open", listViewChronicles.SelectedItem.ListSubItems(9) & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(listViewChronicles.SelectedItem.ListSubItems(9), 3), 1
End Sub



Private Sub listViewCabinets_DblClick()
art_name = listViewCabinets.SelectedItem.ListSubItems(4)
aut_name = listViewCabinets.SelectedItem.ListSubItems(5)
art_date = listViewCabinets.SelectedItem.ListSubItems(6)
file_type = listViewCabinets.SelectedItem.ListSubItems(8)

ShellExecute hwnd, "open", listViewCabinets.SelectedItem.ListSubItems(9) & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(listViewCabinets.SelectedItem.ListSubItems(9), 3), 1
End Sub

Private Sub listViewDocuments_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Local Error GoTo errorOccur



If textDetails.Text <> varTextDetailsChange Then
    If MsgBox("Do you want to save the changes for comments?" & vbCr & vbLf & "Article Name: " & saveart_name & vbCr & vbLf & "Author: " & saveaut_name & vbCr & vbLf & "Sending Date: " & savedate_format, vbYesNo) = vbYes Then
        Call savecomments
    End If
End If


If ListViewDocuments.SelectedItem Is Nothing Then
Else
    art_path = ListViewDocuments.SelectedItem.ListSubItems(9)
    art_name = ListViewDocuments.SelectedItem.ListSubItems(4)
    aut_name = ListViewDocuments.SelectedItem.ListSubItems(5)
    art_date = ListViewDocuments.SelectedItem.ListSubItems(6)
    file_type = ListViewDocuments.SelectedItem.ListSubItems(8)
    labelStatus = "Status: The selected item is " & art_name
    
    Open art_path & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd" For Input As 1
        i = 1
        Do Until EOF(1)
            Input #1, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
            If i = 1 Then
                textDetails = art_keys
                textScore = review_score
                textAuthor = aut_name
                textYear = art_date
                textVenue = art_conf
                textTitle = art_name
                If review_status = "1" Then
                    opt_unreviewed = True
                ElseIf review_status = "2" Then
                    opt_abstract = True
                ElseIf review_status = "3" Then
                    opt_skim = True
                ElseIf review_status = "4" Then
                    opt_read = True
                End If
            End If
            i = i + 1
        Loop
    Close 1
    
    varTextDetailsChange = textDetails.Text
    savepath = ListViewDocuments.SelectedItem.ListSubItems(9)
    saveart_name = ListViewDocuments.SelectedItem.ListSubItems(4)
    saveaut_name = ListViewDocuments.SelectedItem.ListSubItems(5)
    saveart_date = ListViewDocuments.SelectedItem.ListSubItems(6)
    savefile_type = ListViewDocuments.SelectedItem.ListSubItems(8)
    savedate_format = ListViewDocuments.SelectedItem
End If

Exit Sub
errorOccur:
ListViewDocuments.ListItems.Remove ListViewDocuments.SelectedItem.Index
labelStatus = "Status: The selected item couldn't find and removed from the list..."
tabMain.Tabs(1).Caption = "Documents (" & ListViewDocuments.ListItems.Count & ")"
End Sub

Private Sub listViewChronicles_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Local Error GoTo errorOccur

If textDetails.Text <> varTextDetailsChange Then
    If MsgBox("Do you want to save the changes for comments?" & vbCr & vbLf & "Article Name: " & saveart_name & vbCr & vbLf & "Author: " & saveaut_name & vbCr & vbLf & "Sending Date: " & savedate_format, vbYesNo) = vbYes Then
        Call savecomments
    End If
End If


If listViewChronicles.SelectedItem Is Nothing Then
Else
    art_path = listViewChronicles.SelectedItem.ListSubItems(9)
    art_name = listViewChronicles.SelectedItem.ListSubItems(4)
    aut_name = listViewChronicles.SelectedItem.ListSubItems(5)
    art_date = listViewChronicles.SelectedItem.ListSubItems(6)
    file_type = listViewChronicles.SelectedItem.ListSubItems(8)
    labelStatus = "Status: The selected item is " & art_name
    
    Open art_path & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd" For Input As 1
        i = 1
        Do Until EOF(1)
            Input #1, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
            If i = 1 Then
                textDetails = art_keys
                textScore = review_score
                textAuthor = aut_name
                textYear = art_date
                textVenue = art_conf
                textTitle = art_name
                If review_status = "1" Then
                    opt_unreviewed = True
                ElseIf review_status = "2" Then
                    opt_abstract = True
                ElseIf review_status = "3" Then
                    opt_skim = True
                ElseIf review_status = "4" Then
                    opt_read = True
                End If
            End If
            i = i + 1
        Loop
    Close 1
    
    varTextDetailsChange = textDetails.Text
    savepath = listViewChronicles.SelectedItem.ListSubItems(9)
    saveart_name = listViewChronicles.SelectedItem.ListSubItems(4)
    saveaut_name = listViewChronicles.SelectedItem.ListSubItems(5)
    saveart_date = listViewChronicles.SelectedItem.ListSubItems(6)
    savefile_type = listViewChronicles.SelectedItem.ListSubItems(8)
    savedate_format = listViewChronicles.SelectedItem
End If

Exit Sub
errorOccur:
listViewChronicles.ListItems.Remove listViewChronicles.SelectedItem.Index
labelStatus = "Status: The selected item couldn't find and removed from the list..."
tabMain.Tabs(2).Caption = "Chronicles (" & listViewChronicles.ListItems.Count & ")"
End Sub




Private Sub listViewCabinets_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Local Error GoTo errorOccur


If textDetails.Text <> varTextDetailsChange Then
    If MsgBox("Do you want to save the changes for comments?" & vbCr & vbLf & "Article Name: " & saveart_name & vbCr & vbLf & "Author: " & saveaut_name & vbCr & vbLf & "Sending Date: " & savedate_format, vbYesNo) = vbYes Then
        Call savecomments
    End If
End If


If listViewCabinets.SelectedItem Is Nothing Then
Else
    art_path = listViewCabinets.SelectedItem.ListSubItems(9)
    art_name = listViewCabinets.SelectedItem.ListSubItems(4)
    aut_name = listViewCabinets.SelectedItem.ListSubItems(5)
    art_date = listViewCabinets.SelectedItem.ListSubItems(6)
    file_type = listViewCabinets.SelectedItem.ListSubItems(8)
    labelStatus = "Status: The selected item is " & art_name
    
    Open art_path & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd" For Input As 1
        i = 1
        Do Until EOF(1)
            Input #1, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, directoryy
            If i = 1 Then
                textDetails = art_keys
                textScore = review_score
                textAuthor = aut_name
                textYear = art_date
                textVenue = art_conf
                textTitle = art_name
                If review_status = "1" Then
                    opt_unreviewed = True
                ElseIf review_status = "2" Then
                    opt_abstract = True
                ElseIf review_status = "3" Then
                    opt_skim = True
                ElseIf review_status = "4" Then
                    opt_read = True
                End If
            End If
            i = i + 1
        Loop
    Close 1
    
    varTextDetailsChange = textDetails.Text
    savepath = listViewCabinets.SelectedItem.ListSubItems(9)
    saveart_name = listViewCabinets.SelectedItem.ListSubItems(4)
    saveaut_name = listViewCabinets.SelectedItem.ListSubItems(5)
    saveart_date = listViewCabinets.SelectedItem.ListSubItems(6)
    savefile_type = listViewCabinets.SelectedItem.ListSubItems(8)
    savedate_format = listViewCabinets.SelectedItem
End If

Exit Sub
errorOccur:
listViewCabinets.ListItems.Remove listViewCabinets.SelectedItem.Index
labelStatus = "Status: The selected item couldn't find and removed from the list..."
tabMain.Tabs(3).Caption = "Cabinets (" & listViewCabinets.ListItems.Count & ")"
End Sub


Private Sub listViewDocuments_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    art_name = ListViewDocuments.SelectedItem.ListSubItems(4)
    aut_name = ListViewDocuments.SelectedItem.ListSubItems(5)
    art_date = ListViewDocuments.SelectedItem.ListSubItems(6)
    file_type = ListViewDocuments.SelectedItem.ListSubItems(8)
    
    ShellExecute hwnd, "open", ListViewDocuments.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(ListViewDocuments.SelectedItem.ListSubItems(9), 3), 1
End If

End Sub

Private Sub listViewChronicles_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    art_name = listViewChronicles.SelectedItem.ListSubItems(4)
    aut_name = listViewChronicles.SelectedItem.ListSubItems(5)
    art_date = listViewChronicles.SelectedItem.ListSubItems(6)
    file_type = listViewChronicles.SelectedItem.ListSubItems(8)
    
    ShellExecute hwnd, "open", listViewChronicles.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(listViewChronicles.SelectedItem.ListSubItems(9), 3), 1
End If
End Sub

Private Sub listViewCabinets_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    art_name = listViewCabinets.SelectedItem.ListSubItems(4)
    aut_name = listViewCabinets.SelectedItem.ListSubItems(5)
    art_date = listViewCabinets.SelectedItem.ListSubItems(6)
    file_type = listViewCabinets.SelectedItem.ListSubItems(8)
    
    ShellExecute hwnd, "open", listViewCabinets.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(listViewCabinets.SelectedItem.ListSubItems(9), 3), 1
End If
End Sub


Private Sub listViewChronicles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
listViewChronicles.Sorted = True
listViewChronicles.SortKey = ColumnHeader.Index - 1
If listViewChronicles.SortOrder = lvwDescending Then
    listViewChronicles.SortOrder = lvwAscending
Else
    listViewChronicles.SortOrder = lvwDescending
    l = 0
End If
End Sub


Private Sub listViewCabinets_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
listViewCabinets.Sorted = True
listViewCabinets.SortKey = ColumnHeader.Index - 1
If listViewCabinets.SortOrder = lvwDescending Then
    listViewCabinets.SortOrder = lvwAscending
Else
    listViewCabinets.SortOrder = lvwDescending
End If
End Sub




Private Sub tabMain_Click()
If (tabMain.SelectedItem.Index) = 1 Then
    Call doc_vis
ElseIf (tabMain.SelectedItem.Index) = 2 Then
    Call cro_vis
ElseIf (tabMain.SelectedItem.Index) = 3 Then
    Call sof_vis
End If


End Sub



Private Sub textSearchInChronicles_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call new_ordertree(True, False, False)
End If
End Sub

Private Sub textSearchInCabinets_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call new_ordertree(False, True, False)
End If
End Sub

Private Sub textSearchInDocuments_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call new_ordertree(False, False, True)
End If
End Sub


Private Sub textDetails_Change()
textDetails.Text = retextQ(textDetails.Text)
End Sub

Public Function retextQ(strr)
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, """")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    retextQ = strr
End Function




Private Sub ResizeControls(frm As Form)
Dim i As Integer
'   Get ratio of initial form size to current form size

'Loop though all the objects on the form
'Based on the upper bound of the # of controls
For i = 0 To UBound(list)
    'Grad each control individually
    For Each curr_obj In frm
        'Check to make sure its the right control
        If curr_obj.Tag = "unsprt" Then
            GoTo Unsupported
        End If
        If curr_obj.TabIndex = list(i).Index Then
            'Then resize the control
             With curr_obj
                x_size = frm.height / list(i).iHeight
                y_size = frm.width / list(i).iWidth

                .Left = list(i).Left * y_size
                .width = list(i).width * y_size
                
                .Top = list(i).Top * x_size
                On Local Error GoTo Hata:
                .height = list(i).height * x_size

             End With
        End If
Hata:    'Get the next control
Unsupported:
    Next curr_obj
Next i
End Sub


Private Sub GetLocation(frm As Form)
On Error Resume Next

Dim i As Integer
For Each curr_obj In frm
    ReDim Preserve list(i)
    If curr_obj.Tag = "unsprt" Then
        GoTo Unsupported
    End If
    With list(i)
        .Index = curr_obj.TabIndex
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
        .iHeight = frm.height
        .iWidth = frm.width
    End With
    i = i + 1
Unsupported:
Next curr_obj
End Sub






Private Sub SearchForFiles(sRoot As String, listFilePath As ListBox, listFolderPath As ListBox)

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
  
   hFile = FindFirstFile(sRoot & ALL_FILES, WFD)
  
   If hFile <> INVALID_HANDLE_VALUE Then
   
      Do
                  
        'if a folder, and recurse specified, call
        'method again
         If (WFD.dwFileAttributes And vbDirectory) Then
            If Asc(WFD.cFileName) <> vbDot Then

             If fp.bRecurse Then
                  Call SearchForFiles(sRoot & TrimNull(WFD.cFileName) & vbBackslash, listFilePath, listFolderPath)
               End If
            End If
            
         Else
         
           'must be a file..
            If MatchSpec(WFD.cFileName, fp.sFileNameExt) Then
               fp.nCount = fp.nCount + 1
               listFilePath.AddItem sRoot & TrimNull(WFD.cFileName)
               listFolderPath.AddItem sRoot
            End If  'If MatchSpec
      
         End If 'If WFD.dwFileAttributes
      
         fp.nSearched = fp.nSearched + 1
      
      Loop While FindNextFile(hFile, WFD)
   
   End If 'If hFile
  
   Call FindClose(hFile)

End Sub


Private Function removeLastCharacter(str As String) As String

If lstrlen(StrPtr(str)) > 0 Then
   removeLastCharacter = Left$(str, lstrlen(StrPtr(str)) - 1)
Else
   removeLastCharacter = str
End If
      
End Function

Private Function QualifyPath(sPath As String) As String

   If Right$(sPath, 1) <> vbBackslash Then
      QualifyPath = sPath & vbBackslash
   Else
      QualifyPath = sPath
   End If
      
End Function


Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlen(StrPtr(startstr)))
   
End Function


Private Function MatchSpec(sFile As String, sSpec As String) As Boolean

   MatchSpec = PathMatchSpec(StrPtr(sFile), StrPtr(sSpec))
   
End Function

