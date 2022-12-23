VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formAddRecoding 
   Caption         =   "Add a New Record"
   ClientHeight    =   5010
   ClientLeft      =   22545
   ClientTop       =   8175
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   11265
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   6840
      TabIndex        =   11
      Top             =   4200
      Width           =   2052
   End
   Begin VB.Frame Frame5 
      Caption         =   "Starting position for play (in milliseconds)"
      Height          =   612
      Left            =   6720
      TabIndex        =   32
      Top             =   2160
      Width           =   4332
      Begin MSComctlLib.Slider Slider1 
         Height          =   252
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "You can choose a beginning for playing the recording"
         Top             =   240
         Width           =   4092
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   500
         SmallChange     =   100
         TickStyle       =   3
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   492
      Left            =   9720
      TabIndex        =   31
      Top             =   2880
      Width           =   1332
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Recording"
      Enabled         =   0   'False
      Height          =   1032
      Left            =   9050
      TabIndex        =   30
      Top             =   3480
      Width           =   2004
   End
   Begin VB.Frame Frame2 
      Caption         =   "Channels"
      Enabled         =   0   'False
      Height          =   852
      Left            =   6720
      TabIndex        =   27
      Top             =   1200
      Width           =   2052
      Begin VB.OptionButton optStereo 
         Caption         =   "stereo"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optMono 
         Caption         =   "mono"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resolution"
      Enabled         =   0   'False
      Height          =   852
      Left            =   8880
      TabIndex        =   24
      Top             =   1200
      Width           =   2172
      Begin VB.OptionButton opt16bits 
         Caption         =   "16 bits"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opt8bits 
         Caption         =   "8 bits"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample rate (Hz)"
      Enabled         =   0   'False
      Height          =   2172
      Left            =   5160
      TabIndex        =   18
      Top             =   1200
      Width           =   1452
      Begin VB.OptionButton optRate44100 
         Caption         =   "44100"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optRate22050 
         Caption         =   "22050"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optRate11025 
         Caption         =   "11025"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optRate8000 
         Caption         =   "8000"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optRate6000 
         Caption         =   "6000"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   492
      Left            =   8280
      TabIndex        =   17
      Top             =   2880
      Width           =   1332
   End
   Begin VB.OptionButton opt_skim 
      Caption         =   "Skim"
      Height          =   312
      Left            =   8040
      TabIndex        =   14
      Top             =   3480
      Width           =   684
   End
   Begin VB.OptionButton opt_unreviewed 
      Caption         =   "Unreviewed"
      Height          =   312
      Left            =   6840
      TabIndex        =   12
      Top             =   3480
      Width           =   1164
   End
   Begin VB.OptionButton opt_read 
      Caption         =   "Read"
      Height          =   192
      Left            =   8040
      TabIndex        =   13
      Top             =   3840
      Width           =   1284
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   4212
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "AR.frx":0000
      Top             =   360
      Width           =   5052
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "Owner"
      Top             =   720
      Width           =   4212
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Width           =   4212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   492
      Left            =   6720
      TabIndex        =   0
      Top             =   2880
      Width           =   1452
   End
   Begin VB.OptionButton opt_abstract 
      Caption         =   "Abstract"
      Height          =   192
      Left            =   6840
      TabIndex        =   15
      Top             =   3840
      Width           =   1164
   End
   Begin VB.Label Label10 
      Caption         =   "Recording Status:"
      Height          =   252
      Left            =   5160
      TabIndex        =   16
      Top             =   3480
      Width           =   2052
   End
   Begin VB.Label CurrentDir 
      Caption         =   "Label8"
      Height          =   372
      Left            =   9000
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status: Ready"
      Height          =   312
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   11052
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "About Recording:"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1332
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Recording Owner"
      Height          =   252
      Left            =   5280
      TabIndex        =   6
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Recording Date:"
      Height          =   252
      Index           =   1
      Left            =   7200
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Recording Caption:"
      Height          =   252
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label Label8 
      Caption         =   "Importance (0-100):"
      Height          =   252
      Left            =   5160
      TabIndex        =   10
      Top             =   4200
      Width           =   2052
   End
End
Attribute VB_Name = "formAddRecoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rate_Rec, Resolution, Channels
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


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

Private Sub Command1_Click()
Dim i As Integer
'Close any MCI operations from previous VB programs
i = mciSendString("close all", 0&, 0, 0)
 
'Open a new WAV with MCI Command...
i = mciSendString("open new type waveaudio alias capture", 0&, 0, 0)
'Samples Per Second that are supported:
'11025   low quality
'22050   medium quality
'44100 high quality (CD music quality)
 
 
'Bits per sample is 16 or 8


'Channels are 1 (mono) or 2 (stereo)

i = mciSendString("set capture channels " & Channels, 0&, 0, 0) ' 2 channels for stereo
  
   'start at begining
i = mciSendString("seek capture to start", 0&, 0, 0) 'Always start at the beginning

i = mciSendString("set capture samplespersec " & Rate_Rec, 0&, 0, 0) 'CD Quality

i = mciSendString("set capture bitspersample " & Resolution, 0&, 0, 0)  '16 bits for better sound

i = mciSendString("record capture", 0&, 0, 0)  'Start the recording

Command2.Enabled = True   'Enable the STOP BUTTON
Command3.Enabled = False  'Disable the "PLAY" button
Command4.Enabled = False   'Disable the "SAVE AS" button
Command1.Enabled = False

End Sub

Public Function makedir(loc As String)
On Local Error GoTo Hata

MkDir (loc)
Hata:
End Function
Public Function status(sta As String)

Label4 = Format(Now, "Long Time") & " Status: " & sta

End Function






Private Sub Command2_Click()
Dim i As Integer
i = mciSendString("stop capture", 0&, 0, 0)

Command4.Enabled = True  'Enable the "SAVE AS" button
Command3.Enabled = True  'Enable the "PLAY" button
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Dim i As Integer
i = mciSendString("play capture from 0", 0&, 0, 0)
End Sub

Private Sub Command4_Click()
On Local Error GoTo hata1

art_name = Text2.Text
from_name = Text4.Text
aut_name = Text4.Text
art_date = Mid$(Year(Now), 1, Len(Year(Now)))
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



If art_name = "" Then
    strr = "Recording caption is empty."
    GoTo hata1:
    End If
    
If aut_name = "" Then
    strr = "author name section is empty."
    GoTo hata1:
    End If
    
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



makedir (CurrentDir & "\documents\" & date_format)
makedir (CurrentDir & "\documents\" & date_format & "\" & from_name)

ChDrive Left$(CurrentDir & "\documents\" & date_format & "\" & from_name & "\", 1)
ChDir (CurrentDir & "\documents\" & date_format & "\" & from_name & "\")

Dim zz As Integer
zz = mciSendString("save capture " & """" & CurrentDir & "\documents\" & date_format & "\" & from_name & "\" & art_name & "_" & aut_name & "_" & art_date & "." & "wav" & """", 0&, 0, 0) 'MCI command to save the WAV file


Open art_name & "_" & aut_name & "_" & art_date & "_" & "wav" & ".afd" For Output As 1
    Write #1, from_name, date_format, art_name, aut_name, "", art_date, art_keys, "wav", review_score, review_status, "documents"
Close

status (Text2.Text & " is added as record")
formTracker.IsAFDpathsForDocumentsUpdated = 0
ChDrive Left$(CurrentDir, 1)
ChDir CurrentDir
Exit Sub
hata1:
status ("Status: File couldn't be added because " & strr)




End Sub

Private Sub Form_Load()
GetLocation formAddRecoding
CurrentDir = formTracker.labelCurrentDir
opt_unreviewed = True

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
ResizeControls formAddRecoding
End Sub


Private Sub Text1_Change()

Text1.Text = retext(Text1.Text)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
'Close any MCI operations from previous VB programs
i = mciSendString("close all", 0&, 0, 0)
End Sub

Private Sub optRate11025_Click()
    Rate_Rec = "11025"
End Sub

Private Sub optRate44100_Click()
    Rate_Rec = "44100"
End Sub


Private Sub optRate22050_Click()
    Rate_Rec = "22050"
End Sub


Private Sub optRate8000_Click()
    Rate_Rec = "8000"
End Sub

Private Sub optRate6000_Click()
    Rate_Rec = "6000"
End Sub

Private Sub optMono_Click()
    Channels = "1"
End Sub

Private Sub optStereo_Click()
    Channels = "2"
End Sub

Private Sub opt8bits_Click()
    Resolution = "8"
End Sub

Private Sub opt16bits_Click()
    Resolution = "16"
End Sub

Private Sub Text2_Change()
Text2.Text = retext(Text2.Text)
End Sub

Private Sub Text4_Change()
Text4.Text = retext(Text4.Text)
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


