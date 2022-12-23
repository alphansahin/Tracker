VERSION 5.00
Begin VB.Form formAddNote 
   Caption         =   "Add a New Note"
   ClientHeight    =   4665
   ClientLeft      =   18585
   ClientTop       =   8325
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   15210
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opt_skim 
      Caption         =   "Skim"
      Height          =   255
      Left            =   12360
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.OptionButton opt_unreviewed 
      Caption         =   "Unreviewed"
      Height          =   255
      Left            =   11040
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton opt_read 
      Caption         =   "Read"
      Height          =   255
      Left            =   12360
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   3852
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "AN.frx":0000
      Top             =   360
      Width           =   9492
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10920
      TabIndex        =   2
      Text            =   "Owner"
      Top             =   720
      Width           =   4212
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10920
      TabIndex        =   1
      Top             =   360
      Width           =   4212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Note"
      Height          =   852
      Left            =   13200
      TabIndex        =   0
      Top             =   1080
      Width           =   1932
   End
   Begin VB.OptionButton opt_abstract 
      Caption         =   "Abstract"
      Height          =   255
      Left            =   11040
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Date: "
      Height          =   252
      Left            =   10800
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label10 
      Caption         =   "Note Status:"
      Height          =   252
      Left            =   9840
      TabIndex        =   15
      Top             =   1080
      Width           =   1572
   End
   Begin VB.Label CurrentDir 
      Caption         =   "Label8"
      Height          =   372
      Left            =   12960
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status: Ready"
      Height          =   192
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   16008
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "About Note:"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1332
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Note Owner"
      Height          =   252
      Left            =   9840
      TabIndex        =   5
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Note Caption:"
      Height          =   252
      Left            =   9840
      TabIndex        =   4
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label Label8 
      Caption         =   "Importance (0-100):"
      Height          =   252
      Left            =   9840
      TabIndex        =   9
      Top             =   1680
      Width           =   1572
   End
End
Attribute VB_Name = "formAddNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
On Local Error GoTo hata1

art_name = Text2.Text
from_name = Text4.Text
aut_name = Text4.Text
art_date = Mid$(Year(Now), 1, Len(str(Year(Now))))
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
    strr = "note caption is empty."
    GoTo hata1:
    End If
    
If aut_name = "" Then
    strr = "author name section is empty."
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


makedir (CurrentDir & "\documents\" & date_format)
makedir (CurrentDir & "\documents\" & date_format & "\" & from_name)

ChDrive Left$(CurrentDir & "\documents\" & date_format & "\" & from_name & "\", 1)
ChDir (CurrentDir & "\documents\" & date_format & "\" & from_name & "\")

Open art_name & "_" & aut_name & "_" & art_date & "." & "txt" For Output As 1
Close

Open art_name & "_" & aut_name & "_" & art_date & "_" & "txt" & ".afd" For Output As 1
    Write #1, from_name, date_format, art_name, aut_name, "", art_date, art_keys, "txt", review_score, review_status, "documents"
Close

status (Text2.Text & " is added as note")

formTracker.IsAFDpathsForDocumentsUpdated = 0
ChDrive Left$(CurrentDir, 1)
ChDir CurrentDir
Exit Sub
hata1:
status ("Status: File couldn't be added because " & strr)
End Sub

Public Function makedir(loc As String)
On Local Error GoTo Hata

MkDir (loc)
Hata:
End Function
Public Function status(sta As String)

Label4 = Format(Now, "Long Time") & " Status: " & sta

End Function



Private Sub Form_Load()
GetLocation formAddNote
CurrentDir = formTracker.labelCurrentDir
opt_read = True

status ("Ready")
Label3.Caption = "Date: " & Date
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
ResizeControls formAddNote
End Sub


Private Sub Text1_Change()

Text1.Text = retext(Text1.Text)

End Sub

Private Sub Label1_Click(Index As Integer)

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


