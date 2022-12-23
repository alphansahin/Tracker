VERSION 5.00
Begin VB.Form formAddCabinet 
   Caption         =   "Add a New Cabinet"
   ClientHeight    =   6075
   ClientLeft      =   17640
   ClientTop       =   7965
   ClientWidth     =   16215
   LinkTopic       =   "Form2"
   ScaleHeight     =   6075
   ScaleWidth      =   16215
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opt_skim 
      Caption         =   "Skim"
      Height          =   255
      Left            =   14160
      TabIndex        =   24
      Top             =   1320
      Width           =   1452
   End
   Begin VB.OptionButton opt_unreviewed 
      Caption         =   "Unreviewed"
      Height          =   255
      Left            =   14160
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton opt_read 
      Caption         =   "Read"
      Height          =   255
      Left            =   14160
      TabIndex        =   23
      Top             =   1560
      Width           =   1452
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9480
      TabIndex        =   21
      Top             =   1920
      Width           =   3132
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   2412
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2640
      Width           =   7812
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12840
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9480
      TabIndex        =   8
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9480
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   4905
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   3972
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   4590
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9480
      TabIndex        =   3
      Top             =   480
      Width           =   6492
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9480
      TabIndex        =   2
      Top             =   1560
      Width           =   3132
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Cabinet"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   15975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   9480
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.OptionButton opt_abstract 
      Caption         =   "Abstract"
      Height          =   255
      Left            =   14160
      TabIndex        =   25
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Review Status:"
      Height          =   252
      Left            =   12960
      TabIndex        =   26
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label CurrentDir 
      Caption         =   "Label8"
      Height          =   372
      Left            =   11280
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status: Ready"
      Height          =   192
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   16008
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "Keywords:"
      Height          =   252
      Left            =   8280
      TabIndex        =   17
      Top             =   2400
      Width           =   852
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Caption         =   "Year:"
      Height          =   252
      Left            =   8280
      TabIndex        =   16
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Label5 
      Caption         =   "Company:"
      Height          =   252
      Left            =   8280
      TabIndex        =   15
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Sending Date:"
      Height          =   252
      Index           =   1
      Left            =   14160
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File Location:"
      Height          =   192
      Left            =   8280
      TabIndex        =   13
      Top             =   120
      Width           =   948
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
      Height          =   252
      Left            =   8280
      TabIndex        =   12
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Source:"
      Height          =   252
      Index           =   0
      Left            =   8280
      TabIndex        =   11
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label Label8 
      Caption         =   "Review Score (0-100):"
      Height          =   612
      Left            =   8280
      TabIndex        =   20
      Top             =   1920
      Width           =   1212
   End
End
Attribute VB_Name = "formAddCabinet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ITEMID
End Type

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Const CSIDL_DESKTOP = &H0
Const CSIDL_INTERNET = &H1
Const CSIDL_PROGRAMS = &H2
Const CSIDL_CONTROLS = &H3
Const CSIDL_PRINTERS = &H4
Const CSIDL_PERSONAL = &H5
Const CSIDL_FAVORITES = &H6
Const CSIDL_STARTUP = &H7
Const CSIDL_RECENT = &H8
Const CSIDL_SENDTO = &H9
Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const CSIDL_COMMON_STARTMENU = &H16
Const CSIDL_COMMON_PROGRAMS = &H17
Const CSIDL_COMMON_STARTUP = &H18
Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Const CSIDL_APPDATA = &H1A
Const CSIDL_PRINTHOOD = &H1B
Const CSIDL_ALTSTARTUP = &H1D
Const CSIDL_COMMON_ALTSTARTUP = &H1E
Const CSIDL_COMMON_FAVORITES = &H1F
Const CSIDL_INTERNET_CACHE = &H20
Const CSIDL_COOKIES = &H21
Const CSIDL_HISTORY = &H22
Const MAX_strPath = 260

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

Public Function GetType(loc)

For i = 1 To Len(loc)
    dot_point = InStr(Len(loc) - i, loc, ".")
    If dot_point = 0 Then
    
    Else
        GetType = Mid$(loc, dot_point + 1, Len(loc))
        Exit Function
    End If
Next i


End Function


Private Function GetSpecialfolder(CSIDL As Long) As String
    Dim lngRetVal As Long
    Dim IDL As ITEMIDLIST
    Dim strPath As String
    'Get the special folder
    lngRetVal = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lngRetVal = 0 Then
        'Create a buffer
        strPath$ = Space$(512)
        'Get the strPath from the IDList
        lngRetVal = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Private Sub Command1_Click()
On Local Error GoTo hata1
    
file_type = GetType(Text3.Text)

file_loc = Text3.Text
from_name = Text1.Text
art_name = Text2.Text
aut_name = Text4.Text
art_date = Text5.Text
art_conf = Text6.Text
art_keys = Text8.Text
review_score = Text7.Text

If opt_unreviewed = True Then
    review_status = "1"
ElseIf opt_abstract = True Then
    review_status = "2"
ElseIf opt_skim = True Then
    review_status = "3"
ElseIf opt_read = True Then
    review_status = "4"
End If

If file_loc = "" Then
    strr = "file isn't selected."
    GoTo hata1:
    End If

If from_name = "" Then
    strr = "from section is empty."
    GoTo hata1:
    End If


If art_name = "" Then
    strr = "article name is empty."
    GoTo hata1:
    End If
    
If aut_name = "" Then
    strr = "author name section is empty."
    GoTo hata1:
    End If
    
If art_date = "" Then
    strr = "article date is empty."
    GoTo hata1:
    End If

sttr = ""


ss = Day(Now)
If Len(ss) = 1 Then
    day_format = "0" & ss
Else
    day_format = ss
End If

ss = Month(Now)
If Len(ss) = 1 Then
    month_format = "0" & ss
Else
    month_format = ss
End If

date_format = Year(Now) & month_format & day_format



makedir (CurrentDir & "\cabinets" & "\" & date_format)
makedir (CurrentDir & "\cabinets" & "\" & date_format & "\" & from_name)

strr = "file access error is occured. Maybe file is in usage."
FileCopy file_loc, CurrentDir & "\cabinets" & "\" & date_format & "\" & from_name & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type
strr = ""
status (File1.FileName & " is added")

ChDrive Left$(CurrentDir & "\cabinets" & "\" & date_format & "\" & from_name & "\", 1)
ChDir (CurrentDir & "\cabinets" & "\" & date_format & "\" & from_name & "\")

Open art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd" For Output As 1
    Write #1, from_name, date_format, art_name, aut_name, art_conf, art_date, art_keys, file_type, review_score, review_status, "cabinets"
Close
Dir1.Refresh
formTracker.IsAFDpathsForCabinetsUpdated = 0
ChDrive Left$(CurrentDir, 1)
ChDir CurrentDir
Exit Sub
hata1:
status ("Status: File couldn't be added" & " because " & strr)
End Sub

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

For i = 1 To Len(loc)
    dot_point = InStr(Len(loc) - i, loc, ".")
    If dot_point = 0 Then
    
    Else
        TypeName = Mid$(loc, dot_point + 1, Len(loc))
        Exit Function
    End If
Next i


End Function


Public Function makedir(loc As String)
On Local Error GoTo Hata

MkDir (loc)
Hata:
End Function
Public Function status(sta As String)

Label4 = Format(Now, "Long Time") & " Status: " & sta

End Function

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text3.Text = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Text3.Text = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub File1_Click()
Text3.Text = Dir1.Path & "\" & File1.FileName

End Sub

Private Sub File1_DblClick()
On Local Error GoTo Hata
ShellExecute hwnd, "open", File1.Path & "\" & File1.FileName, vbNullString, Left$(File1.Path, 3), 1
Exit Sub
Hata:
End Sub

Private Sub Form_Load()
GetLocation formAddCabinet
desktopDir = GetSpecialfolder(CSIDL_DESKTOP)
File1.Path = desktopDir
Dir1.Path = desktopDir
opt_unreviewed = True

CurrentDir = formTracker.labelCurrentDir

status ("Ready")
Label1(1).Caption = "Date: " & Date
End Sub

Public Function retext(strr)
     If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "/")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "\")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, ":")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "?")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, ">")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "<")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "|")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
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
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "?")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "|")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If

    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, "*")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, ".")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    
    If Len(strr) > 0 Then
        error_dot = InStr(1, strr, ",")
        If Len(strr) = error_dot Then
            strr = Mid$(strr, 1, Len(strr) - 1)
        ElseIf error_dot = 1 Then
            strr = Mid$(strr, 2, Len(strr))
        ElseIf error_dot > 0 Then
            strr = Mid$(strr, 1, error_dot - 1) & Mid$(strr, error_dot + 1, Len(strr))
        End If
    End If
    
    
retext = strr
End Function






Private Sub Form_Resize()
ResizeControls formAddCabinet
End Sub


Private Sub Text1_Change()

Text1.Text = retext(Text1.Text)

End Sub

Private Sub Text2_Change()
Text2.Text = retext(Text2.Text)
End Sub

Private Sub Text4_Change()
Text4.Text = retext(Text4.Text)
End Sub

Private Sub Text5_Change()
Text5.Text = retext(Text5.Text)
End Sub

Private Sub Text8_Change()
Text8.Text = retextQ(Text8.Text)
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
        
        For lngIndex = 1 To Len(strr)
            If Asc(Mid$(strr, lngIndex, 1)) > 7 And Asc(Mid$(strr, lngIndex, 1)) < 15 Or _
               Asc(Mid$(strr, lngIndex, 1)) > 30 And Asc(Mid$(strr, lngIndex, 1)) < 34 Or _
               Asc(Mid$(strr, lngIndex, 1)) > 34 And Asc(Mid$(strr, lngIndex, 1)) < 127 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 128 Or Asc(Mid$(strr, lngIndex, 1)) = 130 Or Asc(Mid$(strr, lngIndex, 1)) = 131 Or Asc(Mid$(strr, lngIndex, 1)) = 134 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 177 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 178 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 179 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 180 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 134 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 252 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 246 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 231 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 199 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 214 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 220 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 247 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 175 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 176 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 186 Or _
               Asc(Mid$(strr, lngIndex, 1)) = 185 Or _
               IsNumeric(Mid$(strr, lngIndex, 1)) Then
                strNew = strNew & Mid$(strr, lngIndex, 1)
            End If
        Next
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


