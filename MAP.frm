VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formActiveFolders 
   Caption         =   "Active Folders"
   ClientHeight    =   9990
   ClientLeft      =   2610
   ClientTop       =   1560
   ClientWidth     =   20280
   LinkTopic       =   "Form7"
   ScaleHeight     =   9990
   ScaleWidth      =   20280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   7560
      Width           =   5175
      Begin VB.CommandButton Command11 
         Caption         =   "Down (ctrl + down)"
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Up (ctrl + up)"
         Height          =   495
         Left            =   1800
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh (F5)"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save (ctrl + s)"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Insert branch (shift + ins)"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1600
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Insert keyword (shift + k)"
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Width           =   1600
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Insert subbranch (ins)"
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1600
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Rename branch (F2)"
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   1600
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Delete branch (del)"
         Height          =   495
         Left            =   3480
         TabIndex        =   7
         Top             =   1080
         Width           =   1600
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   4935
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8895
      Left            =   5760
      TabIndex        =   3
      Top             =   480
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   15690
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
   Begin VB.CommandButton Command8 
      Caption         =   "Delete Checked File(s)"
      Height          =   372
      Left            =   13320
      TabIndex        =   2
      Top             =   9480
      Width           =   6825
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Export Checked File(s)"
      Height          =   372
      Left            =   5760
      TabIndex        =   1
      Top             =   9480
      Width           =   7425
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Tag             =   "unsprt"
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3840
      Tag             =   "unsprt"
      Top             =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7092
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5172
      _ExtentX        =   9128
      _ExtentY        =   12515
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TabStrip TabStrip3 
      Height          =   9855
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   17383
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   9855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   17383
      TabMinWidth     =   706
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Folders"
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
End
Attribute VB_Name = "formActiveFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const MAX_PATH = 260
Private treeloaded


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


Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim mfX As Single
Dim mfY As Single

Public k

Private indexLayer0
Private indexLayer1
Private indexLayer2
Private indexLayer3
Private indexLayer4
Private indexKeyword

Private imgx



Private Sub Command1_Click()
     
    loadMap
End Sub


Private Sub deleteItem()
    If Not TreeView1.SelectedItem Is Nothing Then
        TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
    End If
End Sub


Private Sub editItem()
    If Not TreeView1.SelectedItem Is Nothing Then
        TreeView1.StartLabelEdit
    End If
End Sub


Private Sub saveMap()
    On Error Resume Next
    Kill "mapT.txt"
    i = 0
    indexL0 = TreeView1.Nodes.Item(1).FirstSibling.Index
    doneL0 = 0
    Do While doneL0 = 0

            saveMapRec TreeView1.Nodes.Item(indexL0), 0

        If indexL0 = TreeView1.Nodes.Item(indexL0).Next.Index Then
            doneL0 = 1
        Else
            indexL0 = TreeView1.Nodes.Item(indexL0).Next.Index
        End If
    Loop
    Open "mapT.txt" For Append As #1
    Print #1, "end"
    Close
    

    Kill "map.txt"
    Name "mapT.txt" As "map.txt"
    CurDir
    status "The map is saved."
End Sub

Private Sub saveMapRec(ByVal n As Node, spacing)
    Open "mapT.txt" For Append As #1
    If n.Tag = "L0" Then
        If spacing = 1 Then
            Print #1, ""
        End If
    
        Print #1, "<L0>"
        Print #1, n
        
        numberOfKeywords = 0
        Set ChildNode = n.Child
        For kkk = 0 To n.Children - 1
            If ChildNode.Tag = "LK" Then
                numberOfKeywords = ChildNode.Children
                Exit For
            End If
            Set ChildNode = ChildNode.Next
        Next kkk
        
        Print #1, CStr(numberOfKeywords)
        
        If numberOfKeywords <> 0 Then
            Set keywordNode = ChildNode.Child
            For kkk = 0 To numberOfKeywords - 1
                Print #1, keywordNode & ","
                Set keywordNode = keywordNode.Next
            Next kkk
        End If
        
        Print #1, "</L0>"
        
    End If
    
    If n.Tag = "L1" Then
        If spacing = 1 Then
            Print #1, ""
        End If
    
            
        Print #1, "<L1>"
        Print #1, n
        
        numberOfKeywords = 0
        Set ChildNode = n.Child
        For kkk = 0 To n.Children - 1
            If ChildNode.Tag = "LK" Then
                numberOfKeywords = ChildNode.Children
                Exit For
            End If
            Set ChildNode = ChildNode.Next
        Next kkk
        
        Print #1, CStr(numberOfKeywords)
        
        If numberOfKeywords <> 0 Then
            Set keywordNode = ChildNode.Child
            For kkk = 0 To numberOfKeywords - 1
                Print #1, keywordNode & ","
                Set keywordNode = keywordNode.Next
            Next kkk
        End If
        
        Print #1, "</L1>"
        
    End If
    
    If n.Tag = "L2" Then
        If spacing = 1 Then
            Print #1, ""
        End If
    
            
        Print #1, "<L2>"
        Print #1, n
        
        numberOfKeywords = 0
        Set ChildNode = n.Child
        For kkk = 0 To n.Children - 1
            If ChildNode.Tag = "LK" Then
                numberOfKeywords = ChildNode.Children
                Exit For
            End If
            Set ChildNode = ChildNode.Next
        Next kkk
        
        Print #1, CStr(numberOfKeywords)
        
        If numberOfKeywords <> 0 Then
            Set keywordNode = ChildNode.Child
            For kkk = 0 To numberOfKeywords - 1
                Print #1, keywordNode & ","
                Set keywordNode = keywordNode.Next
            Next kkk
        End If
        
        Print #1, "</L2>"
        
    End If
    
    
    If n.Tag = "L3" Then
        If spacing = 1 Then
            Print #1, ""
        End If
    
            
        Print #1, "<L3>"
        Print #1, n
        
        numberOfKeywords = 0
        Set ChildNode = n.Child
        For kkk = 0 To n.Children - 1
            If ChildNode.Tag = "LK" Then
                numberOfKeywords = ChildNode.Children
                Exit For
            End If
            Set ChildNode = ChildNode.Next
        Next kkk
        
        Print #1, CStr(numberOfKeywords)
        
        If numberOfKeywords <> 0 Then
            Set keywordNode = ChildNode.Child
            For kkk = 0 To numberOfKeywords - 1
                Print #1, keywordNode & ","
                Set keywordNode = keywordNode.Next
            Next kkk
        End If
        
        Print #1, "</L3>"

    End If
    
    
    If n.Tag = "L4" Then
        If spacing = 1 Then
            Print #1, ""
        End If
    
    
        Print #1, "<L4>"
        Print #1, n
        
        numberOfKeywords = 0
        Set ChildNode = n.Child
        For kkk = 0 To n.Children - 1
            If ChildNode.Tag = "LK" Then
                numberOfKeywords = ChildNode.Children
                Exit For
            End If
            Set ChildNode = ChildNode.Next
        Next kkk
        
        Print #1, CStr(numberOfKeywords)
        
        If numberOfKeywords <> 0 Then
            Set keywordNode = ChildNode.Child
            For kkk = 0 To numberOfKeywords - 1
                Print #1, keywordNode & ","
                Set keywordNode = keywordNode.Next
            Next kkk
        End If
        
        Print #1, "</L4>"
        

    End If
    Close
    
    
    Dim aNode As Node
    Set aNode = n.Child
    For i = 1 To n.Children
        saveMapRec aNode, 1
        Set aNode = aNode.Next
    Next
    
End Sub




Private Sub Command10_Click()

MoveNode TreeView1, TreeView1.SelectedItem, True
Call saveExpansionInformation
End Sub

Private Sub Command11_Click()
MoveNode TreeView1, TreeView1.SelectedItem, False
Call saveExpansionInformation
End Sub

Private Sub Command2_Click()
    saveMap
End Sub

Private Sub Command3_Click()
insertBranch
End Sub

Private Sub Command4_Click()
insertKeyword
End Sub

Private Sub Command5_Click()
insertSubBranch
End Sub

Private Sub Command6_Click()
editItem
End Sub

Private Sub Command7_Click()
deleteItem
End Sub

Public Sub insertSubBranch()
    If TreeView1.SelectedItem Is Nothing Then
         MsgBox "No branch is selected, hence a subbranch cannot be created. In order to create a new branch press CTRL + INS."
         Exit Sub
     End If
    
     
     If TreeView1.SelectedItem.Tag = "L4" Then
         MsgBox "A new layer cannot be added. Up to five layer is supported."
         Exit Sub
     ElseIf Mid(TreeView1.SelectedItem.Tag, 1, 2) = "LK" Then
         If TreeView1.SelectedItem.Parent.Tag = "L0" Then
             indexLayer1 = indexLayer1 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Key, tvwChild, "L1" & indexLayer1, "New subbranch (L1)", "L1")
             insNode.Tag = "L1"
         ElseIf TreeView1.SelectedItem.Parent.Tag = "L1" Then
             indexLayer2 = indexLayer2 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Key, tvwChild, "L2" & indexLayer2, "New subbranch (L2)", "L2")
             insNode.Tag = "L2"
         ElseIf TreeView1.SelectedItem.Parent.Tag = "L2" Then
             indexLayer3 = indexLayer3 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Key, tvwChild, "L3" & indexLayer3, "New subbranch (L3)", "L3")
             insNode.Tag = "L3"
         ElseIf TreeView1.SelectedItem.Parent.Tag = "L3" Then
             indexLayer4 = indexLayer4 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Key, tvwChild, "L4" & indexLayer4, "New subbranch (L4)", "L4")
             insNode.Tag = "L4"
         ElseIf TreeView1.SelectedItem.Parent.Tag = "L4" Then
             MsgBox "Unofortunately, a new layer cannot be added. Up to five layer is supported."
             Exit Sub
         End If
         
         
         insNode.Selected = True
         insNode.EnsureVisible
         TreeView1.StartLabelEdit
     ElseIf TreeView1.SelectedItem.Tag = "K" Then
         If TreeView1.SelectedItem.Parent.Parent.Tag = "L0" Then
             indexLayer1 = indexLayer1 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Parent.Key, tvwChild, "L1" & indexLayer1, "New subbranch (L1)", "L1")
             insNode.Tag = "L1"
         ElseIf TreeView1.SelectedItem.Parent.Parent.Tag = "L1" Then
             indexLayer2 = indexLayer2 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Parent.Key, tvwChild, "L2" & indexLayer2, "New subbranch (L2)", "L2")
             insNode.Tag = "L2"
         ElseIf TreeView1.SelectedItem.Parent.Parent.Tag = "L2" Then
             indexLayer3 = indexLayer3 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Parent.Key, tvwChild, "L3" & indexLayer3, "New subbranch (L3)", "L3")
             insNode.Tag = "L3"
         ElseIf TreeView1.SelectedItem.Parent.Parent.Tag = "L3" Then
             indexLayer4 = indexLayer4 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Parent.Key, tvwChild, "L4" & indexLayer4, "New subbranch (L4)", "L4")
             insNode.Tag = "L4"
         ElseIf TreeView1.SelectedItem.Parent.Parent.Tag = "L4" Then
             MsgBox "Unofortunately, a new layer cannot be added. Up to five layer is supported."
             Exit Sub
         End If
         
         
         insNode.Selected = True
         insNode.EnsureVisible
         TreeView1.StartLabelEdit
     Else
         If TreeView1.SelectedItem.Tag = "L0" Then
             indexLayer1 = indexLayer1 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "L1" & indexLayer1, "New subbranch (L1)", "L1")
             insNode.Tag = "L1"
         ElseIf TreeView1.SelectedItem.Tag = "L1" Then
             indexLayer2 = indexLayer2 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "L2" & indexLayer2, "New subbranch (L2)", "L2")
             insNode.Tag = "L2"
         ElseIf TreeView1.SelectedItem.Tag = "L2" Then
             indexLayer3 = indexLayer3 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "L3" & indexLayer3, "New subbranch (L3)", "L3")
             insNode.Tag = "L3"
         ElseIf TreeView1.SelectedItem.Tag = "L3" Then
             indexLayer4 = indexLayer4 + 1
             Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "L4" & indexLayer4, "New subbranch (L4)", "L4")
             insNode.Tag = "L4"
         End If
         
         
         insNode.Selected = True
         insNode.EnsureVisible
         TreeView1.StartLabelEdit
     End If
End Sub

Private Sub Command8_Click()
   Dim FSO As New FileSystemObject
   For i = ListView1.ListItems.Count To 1 Step -1
        Set ListView1.SelectedItem = ListView1.ListItems(i)
        If ListView1.SelectedItem.Checked = True Then
        
            date_format = ListView1.SelectedItem
            art_name = ListView1.SelectedItem.ListSubItems(4)
            aut_name = ListView1.SelectedItem.ListSubItems(5)
            art_date = ListView1.SelectedItem.ListSubItems(6)
            file_type = ListView1.SelectedItem.ListSubItems(8)
        
            
            Set fld = FSO.GetFolder(ListView1.SelectedItem.ListSubItems(9))
                
            If fld.Files.Count = 2 Then
                Kill ListView1.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
                Kill ListView1.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
                
                FSO.DeleteFolder fld, True
                
                Set pfld = FSO.GetFolder(formTracker.labelCurrentDir & "\documents\" & date_format & "\")
                
                If pfld.SubFolders.Count = 0 Then
                    FSO.DeleteFolder pfld, True
                Else
                    
                End If
            Else
                Kill ListView1.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "_" & file_type & ".afd"
                Kill ListView1.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            End If
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
            formTracker.IsAFDpathsForDocumentsUpdated = 0
        End If
    Next i
    TabStrip3.Tabs(1).Caption = TreeView1.SelectedItem & " (" & ListView1.ListItems.Count & ")"
    'Call refreshTrackerScreen
End Sub



Private Sub Command9_Click()
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
   
   
   
    For i = 1 To ListView1.ListItems.Count
        Set ListView1.SelectedItem = ListView1.ListItems(i)
        If ListView1.SelectedItem.Checked = True Then
            date_format = ListView1.SelectedItem
            art_name = ListView1.SelectedItem.ListSubItems(4)
            aut_name = ListView1.SelectedItem.ListSubItems(5)
            art_date = ListView1.SelectedItem.ListSubItems(6)
            file_type = ListView1.SelectedItem.ListSubItems(8)
        
            
            SourceFile = ListView1.SelectedItem.ListSubItems(9) & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            DestinationFile = sBuffer & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type
            FileCopy SourceFile, DestinationFile
        End If
    Next i
    
   
End If
MsgBox "File(s) are exported. Path:" & sBuffer

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyS Then
        If vbCtrlMask Then
            saveMap
        End If
    End If
    
    If KeyCode = vbKeyF5 Then
        loadMap
    End If
        
End Sub


Private Sub Form_Load()
    treeloaded = 0
    GetLocation formActiveFolders
    
     ' Add images to ListImages collection.

     
    ImageList1.ListImages.Add , "L0", LoadPicture(CurDir & "\icons\L0.bmp")
    ImageList1.ListImages.Add , "L1", LoadPicture(CurDir & "\icons\L1.bmp")
    ImageList1.ListImages.Add , "L2", LoadPicture(CurDir & "\icons\L2.bmp")
    ImageList1.ListImages.Add , "L3", LoadPicture(CurDir & "\icons\L3.bmp")
    ImageList1.ListImages.Add , "L4", LoadPicture(CurDir & "\icons\L4.bmp")
    ImageList1.ListImages.Add , "LK", LoadPicture(CurDir & "\icons\LK.bmp")
    ImageList1.ListImages.Add , "K", LoadPicture(CurDir & "\icons\K.bmp")
    Set TreeView1.ImageList = ImageList1
    loadMap

    
    Dim n As Integer
    Timer1.Enabled = False
    Timer1.Interval = 20
         
    TreeView1.OLEDragMode = ccOLEDragAutomatic
    TreeView1.OLEDropMode = ccOLEDropManual
    
    
    'Me.WindowState = vbMaximized
         
    Command3.Caption = "Insert branch" & vbCrLf & "(ctrl + ins)"
    Command5.Caption = "Insert subbranch" & vbCrLf & "(ins)"
    Command4.Caption = "Insert keyword" & vbCrLf & "(ctrl + k)"
         
    status "The map is active."
    treeloaded = 1
End Sub

Private Sub Form_Resize()
    ResizeControls formActiveFolders
End Sub

Public Function status(sta As String)

Label1 = Format(Now, "Long Time") & " Status: " & sta

End Function

Private Sub Label3_Click()

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.Sorted = True
ListView1.SortKey = ColumnHeader.Index - 1
If k = 0 Then
    ListView1.SortOrder = lvwAscending
    k = 1
Else
    ListView1.SortOrder = lvwDescending
    k = 0
End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    art_name = ListView1.SelectedItem.ListSubItems(4)
    aut_name = ListView1.SelectedItem.ListSubItems(5)
    art_date = ListView1.SelectedItem.ListSubItems(6)
    file_type = ListView1.SelectedItem.ListSubItems(8)
    
    ShellExecute hwnd, "open", ListView1.SelectedItem.ListSubItems(9) & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(ListView1.SelectedItem.ListSubItems(9), 3), 1
End If
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
art_name = ListView1.SelectedItem.ListSubItems(4)
aut_name = ListView1.SelectedItem.ListSubItems(5)
art_date = ListView1.SelectedItem.ListSubItems(6)
file_type = ListView1.SelectedItem.ListSubItems(8)

ShellExecute hwnd, "open", ListView1.SelectedItem.ListSubItems(9) & "\" & art_name & "_" & aut_name & "_" & art_date & "." & file_type, vbNullString, Left$(ListView1.SelectedItem.ListSubItems(9), 3), 1
End Sub

Private Sub Timer1_Timer()
    Set TreeView1.DropHighlight = TreeView1.HitTest(mfX, mfY)
    If m_iScrollDir = -1 Then
        SendMessage TreeView1.hwnd, 277&, 0&, vbNull
    Else
        SendMessage TreeView1.hwnd, 277&, 1&, vbNull
    End If
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Trim(NewString) = "" Then
        TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
    Else
        Dim Repeat As Boolean
        Repeat = True
        While Repeat
            On Error Resume Next
            TreeView1.SelectedItem.Key = "K" & 1 + Int(Rnd() * 10000000) ' Trim(TreeView1.SelectedItem.FullPath) & NewString
            If Err.Number = 0 Then Repeat = False
        Wend
        If Not TreeView1.SelectedItem.Parent Is Nothing Then
            TreeView1.SelectedItem.Parent.Selected = True
        End If
    End If
End Sub

Private Sub TreeView1_Click()
    Call listUpdate
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
If treeloaded = 1 Then
    Call saveExpansionInformation
End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
If treeloaded = 1 Then
    Call saveExpansionInformation
End If
End Sub

Private Sub TreeView1_Keydown(KeyCode As Integer, Shift As Integer)
    Dim insNode As Node
    If KeyCode = vbKeyF2 Then
        editItem
    End If
    If KeyCode = vbKeyDelete Then
        deleteItem
    End If
    If KeyCode = vbKeyS Then
        If Shift = vbCtrlMask Then
            saveMap
        End If
    End If
    
    If KeyCode = vbKeyK Then
        If vbCtrlMask Then
            insertKeyword
        End If
    End If
        
    
    If KeyCode = vbKeyInsert Then
        If Shift = 0 Then
            insertSubBranch
        Else
            If Shift = vbCtrlMask Then
                insertBranch
            End If
        End If
    End If
    
    If KeyCode = vbKeyF5 Then
        loadMap
    End If
    
    If KeyCode = vbKeyReturn Then
        listUpdate
    End If
    
    If KeyCode = vbKeyUp Then
        If Shift = vbCtrlMask Then
            MoveNode TreeView1, TreeView1.SelectedItem, True
        End If
    End If
    
    If KeyCode = vbKeyDown Then
        If Shift = vbCtrlMask Then
            MoveNode TreeView1, TreeView1.SelectedItem, False
        End If
    End If
    
End Sub

Public Sub insertBranch()
    indexLayer0 = indexLayer0 + 1
    Set insNode = TreeView1.Nodes.Add(, , "L0" & indexLayer0, "New branch (L0)", "L0")
    
    insNode.Selected = True
    TreeView1.StartLabelEdit
    insNode.Tag = "L0"
End Sub

Public Sub insertKeyword()
    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "Select the child node's parent, then press the Ins key to insert child node"
        Exit Sub
    End If

    indexKeyword = indexKeyword + 1
    
    If TreeView1.SelectedItem.Tag = "K" Then
        Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Parent.Index, tvwChild, "K" & indexKeyword, "", "K")
    ElseIf Mid(TreeView1.SelectedItem.Tag, 1, 2) = "LK" Then
        Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "K" & indexKeyword, "", "K")
    Else
        On Error Resume Next
        Set insNode = TreeView1.Nodes.Item("LK" & TreeView1.SelectedItem.Key)
        If Err.Number = 35601 Then
            Set insNode = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "LK" & TreeView1.SelectedItem.Key, "Keywords", "LK")
            insNode.Tag = "LK"
        End If
            
        Set insNode = TreeView1.Nodes.Add(insNode.Key, tvwChild, "K" & indexKeyword, "", "K")
    End If
    insNode.Selected = True
    insNode.Tag = "K"
    
    insNode.Text = "Type keywords (separeted by ','...)"
    insNode.EnsureVisible
    TreeView1.StartLabelEdit

End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.TreeView1.SelectedItem = Me.TreeView1.HitTest(x, y)
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sTemp As String
Dim oNode As Node
Dim oDragNode As Node
  
    Set oNode = Me.TreeView1.HitTest(x, y)
     If Data.GetFormat(vbCFText) Then
        sTemp = Data.GetData(vbCFText)
        'MsgBox sTemp
        Set oDragNode = TreeView1.Nodes(sTemp)
        On Error Resume Next
        If oDragNode.Tag <> "L0" Then
            If (oNode.Tag = oDragNode.Parent.Tag) Then
                Set oDragNode.Parent = oNode
     
                If Err.Number = 35614 Then
                    'MsgBox "Can't create circular relations"
                    On Error GoTo 0
                End If
                Set TreeView1.DropHighlight = Nothing
            End If
        End If
        

        
    End If
End Sub

Private Sub TreeView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    With TreeView1
        If State = vbLeave Then
            Set .DropHighlight = Nothing
        Else
            .DropHighlight = .HitTest(x, y)
        End If
    End With
    mfX = x
    mfY = y
    If y > 0 And y < 100 Then
        m_iScrollDir = -1
        Timer1.Enabled = True
    ElseIf y > (TreeView1.height - 100) And y < TreeView1.height Then
        m_iScrollDir = 1
        Timer1.Enabled = True
        Else
            Timer1.Enabled = False
    End If
End Sub

Private Sub TreeView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  Dim oDragNode As MSComctlLib.Node
  Data.Clear
  If Not Me.TreeView1.SelectedItem Is Nothing Then
    Data.SetData Me.TreeView1.SelectedItem.Key, vbCFText
  End If
End Sub
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

Private Sub listUpdate()
    DoEvents
    If TreeView1.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    
    Dim orArray() As String
    If TreeView1.SelectedItem.Tag <> "K" And TreeView1.SelectedItem.Tag <> "LK" Then
        status "Retrieving the information..."
        
        TabStrip3.Tabs(1).Caption = "Retrieving the information..."
        On Error Resume Next
        ListView1.ListItems.Clear
        Set keywordNode = TreeView1.Nodes.Item("LK" & TreeView1.SelectedItem.Key)
        Set insNode = TreeView1.Nodes.Item("LK" & TreeView1.SelectedItem.Key)
        If Err.Number = 35601 Then
            TabStrip3.Tabs(1).Caption = TreeView1.SelectedItem & " (" & ListView1.ListItems.Count & ")"
            status "The search is completed."
            Exit Sub
        End If
        
        Set ChildNode = keywordNode.Child
        
        ReDim orArray(0 To keywordNode.Children - 1) As String
        
        For kkk = 0 To keywordNode.Children
            
            orArray(kkk) = ChildNode.Text & ","
          
           Set ChildNode = ChildNode.Next
        Next kkk
        
        
        Call formTracker.activeFolderSearch(orArray)

    
    
        For Each oLVItem In formTracker.ListViewDocuments.ListItems
            With formActiveFolders.ListView1.ListItems.Add(, , oLVItem.Text)
                
                .SubItems(1) = oLVItem.SubItems(1)
                .SubItems(2) = oLVItem.SubItems(2)
                .SubItems(3) = oLVItem.SubItems(3)
                .SubItems(4) = oLVItem.SubItems(4)
                .SubItems(5) = oLVItem.SubItems(5)
                .SubItems(6) = oLVItem.SubItems(6)
                .SubItems(7) = oLVItem.SubItems(7)
                .SubItems(8) = oLVItem.SubItems(8)
                .SubItems(9) = oLVItem.SubItems(9)
            End With
        Next
    
        TabStrip3.Tabs(1).Caption = TreeView1.SelectedItem & " (" & ListView1.ListItems.Count & ")"
        status "The search is completed."
    End If
   'Call refreshTrackerScreen
End Sub

Private Sub refreshTrackerScreen()
    formTracker.cmdUpdateList.Enabled = 0
    formTracker.cmdFileDelete.Enabled = 0
    formTracker.cmdFileExport.Enabled = 0
    
    Call formTracker.new_ordertree(True, False, False)
        
    formTracker.cmdUpdateList.Enabled = 1
    formTracker.cmdFileDelete.Enabled = 1
    formTracker.cmdFileExport.Enabled = 1
End Sub


Private Sub loadMap()
treeloaded = 0
TreeView1.Nodes.Clear
    indexLayer0 = 0
    indexLayer1 = 0
    indexLayer2 = 0
    indexLayer3 = 0
    indexLayer4 = 0
    indexKeyword = 0
    
Set TreeView1.ImageList = ImageList1

 Open CurDir & "\map.txt" For Input As #1

    Do While Not EOF(1)        ' Loop until end of file
        Line Input #1, Text  ' Read line into variable
        Select Case Text
            Case "<L0>"
                indexLayer0 = indexLayer0 + 1
                Line Input #1, Title  ' Read line next line
                TreeView1.Nodes.Add , , "L0" & indexLayer0, Title, "L0"
                TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "L0"
                Line Input #1, numberOfKeywords  ' Read line next line
                numKeywords = CInt(numberOfKeywords)
                
                For i = 1 To numKeywords
                    If i = 1 Then
                        Set insNode = TreeView1.Nodes.Add("L0" & indexLayer0, tvwChild, "LK" & "L0" & indexLayer0, "Keywords", "LK")
                        insNode.Tag = "LK"
                    End If
                
                    indexKeyword = indexKeyword + 1
                    Line Input #1, keywords  ' Read line next line
                    Set insNode = TreeView1.Nodes.Add("LK" & "L0" & indexLayer0, tvwChild, "K" & indexKeyword, Mid(keywords, 1, Len(keywords) - 1), "K")
                    insNode.Tag = "K"
                Next
            Case "<L1>"
                indexLayer1 = indexLayer1 + 1
                Line Input #1, Title  ' Read line next line
                TreeView1.Nodes.Add "L0" & indexLayer0, tvwChild, "L1" & indexLayer1, Title, "L1"
                TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "L1"
                Line Input #1, numberOfKeywords  ' Read line next line
                numKeywords = CInt(numberOfKeywords)
                
                For i = 1 To numKeywords
                    If i = 1 Then
                        Set insNode = TreeView1.Nodes.Add("L1" & indexLayer1, tvwChild, "LK" & "L1" & indexLayer1, "Keywords", "LK")
                        insNode.Tag = "LK"
                    End If
                
                    indexKeyword = indexKeyword + 1
                    Line Input #1, keywords  ' Read line next line
                    Set insNode = TreeView1.Nodes.Add("LK" & "L1" & indexLayer1, tvwChild, "K" & indexKeyword, Mid(keywords, 1, Len(keywords) - 1), "K")
                    insNode.Tag = "K"
                Next
            Case "<L2>"
                indexLayer2 = indexLayer2 + 1
                Line Input #1, Title  ' Read line next line
                TreeView1.Nodes.Add "L1" & indexLayer1, tvwChild, "L2" & indexLayer2, Title, "L2"
                TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "L2"
                Line Input #1, numberOfKeywords  ' Read line next line
                numKeywords = CInt(numberOfKeywords)
                
                For i = 1 To numKeywords
                    If i = 1 Then
                        Set insNode = TreeView1.Nodes.Add("L2" & indexLayer2, tvwChild, "LK" & "L2" & indexLayer2, "Keywords", "LK")
                        insNode.Tag = "LK"
                    End If
                
                    indexKeyword = indexKeyword + 1
                    Line Input #1, keywords  ' Read line next line
                    Set insNode = TreeView1.Nodes.Add("LK" & "L2" & indexLayer2, tvwChild, "K" & indexKeyword, Mid(keywords, 1, Len(keywords) - 1), "K")
                    insNode.Tag = "K"
                Next
            Case "<L3>"
                indexLayer3 = indexLayer3 + 1
                Line Input #1, Title  ' Read line next line
                TreeView1.Nodes.Add "L2" & indexLayer2, tvwChild, "L3" & indexLayer3, Title, "L3"
                TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "L3"
                Line Input #1, numberOfKeywords  ' Read line next line
                numKeywords = CInt(numberOfKeywords)
                
                For i = 1 To numKeywords
                    If i = 1 Then
                        Set insNode = TreeView1.Nodes.Add("L3" & indexLayer3, tvwChild, "LK" & "L3" & indexLayer3, "Keywords", "LK")
                        insNode.Tag = "LK"
                    End If
                
                    indexKeyword = indexKeyword + 1
                    Line Input #1, keywords  ' Read line next line
                    Set insNode = TreeView1.Nodes.Add("LK" & "L3" & indexLayer3, tvwChild, "K" & indexKeyword, Mid(keywords, 1, Len(keywords) - 1), "K")
                    insNode.Tag = "K"
                Next
            Case "<L4>"
                indexLayer4 = indexLayer4 + 1
                Line Input #1, Title  ' Read line next line
                TreeView1.Nodes.Add "L3" & indexLayer3, tvwChild, "L4" & indexLayer4, Title, "L4"
                TreeView1.Nodes.Item(TreeView1.Nodes.Count).Tag = "L4"
                Line Input #1, numberOfKeywords  ' Read line next line
                numKeywords = CInt(numberOfKeywords)
                
                For i = 1 To numKeywords
                    If i = 1 Then
                        Set insNode = TreeView1.Nodes.Add("L4" & indexLayer4, tvwChild, "LK" & "L4" & indexLayer4, "Keywords", "LK")
                        insNode.Tag = "LK"
                    End If
                    
                    indexKeyword = indexKeyword + 1
                    Line Input #1, keywords  ' Read line next line
                    Set insNode = TreeView1.Nodes.Add("LK" & "L4" & indexLayer4, tvwChild, "K" & indexKeyword, Mid(keywords, 1, Len(keywords) - 1), "K")
                    insNode.Tag = "K"
                Next
            Case Else
                     
        End Select
    Loop

    Close #1
    
    Dim n As Node
    
    
    Set FSO = New FileSystemObject
    
    If FSO.FileExists(CurDir & "\expansion.txt") Then
        Open CurDir & "\expansion.txt" For Input As #1
        indexNode = 1
        Do While Not EOF(1)        ' Loop until end of file
                Line Input #1, expandOrCollapse  ' Read line next line
                If expandOrCollapse = "1" Then
                    TreeView1.Nodes.Item(indexNode).Expanded = True
                ElseIf expandOrCollapse = "0" Then
                    TreeView1.Nodes.Item(indexNode).Expanded = False
                End If
                indexNode = indexNode + 1
                
        Loop
        Close
    Else
        For Each n In TreeView1.Nodes
        If n.Tag <> "K" And n.Tag <> "LK" Then
            If n.Children > 1 Then
                n.Expanded = True
            Else
                If n.Child Is Nothing Then
                Else
                    If n.Child.Tag <> "K" And n.Child.Tag <> "LK" Then
                        n.Expanded = True
                    End If
                End If
            End If
        End If
    Next
        
    End If
    
    If TreeView1.Nodes.Count > 0 Then
        TreeView1.Nodes(1).EnsureVisible
    End If
    status "The map is reloaded."
    treeloaded = 1
End Sub




Private Sub saveExpansionInformation()
 Open CurDir & "\expansion.txt" For Output As #1
    Dim n As Node
    For Each n In TreeView1.Nodes
            If n.Expanded = True Then
                Print #1, "1"
            Else
                Print #1, "0"
            End If
    Next
    Close
End Sub



