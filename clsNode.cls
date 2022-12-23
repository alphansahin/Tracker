VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ChildNodes"
Option Explicit
Public Tag As Variant
Public Backcolor As Long
Public Bold As Boolean
Public Checked As Boolean
Public Expanded As Boolean
Public ExpandedImage As Long
Public ForeColor As Long
Public Image As String
Public Key As Variant
Public SelectedImage As Long
Public Sorted As Boolean
Public Text_ As String
Public Visible As Boolean
Private mvarChildNodes As New Collection 'ChildNodes

Public Property Get ChildNodes() As Collection
    Set ChildNodes = mvarChildNodes
End Property





Public Property Set ChildNodes(vData As Collection)
    Set mvarChildNodes = vData
End Property
Public Sub CopyNode(tmpNode As Node)
Dim ChildNode As clsNode
Dim cNode As Node
Backcolor = tmpNode.Backcolor
Bold = tmpNode.Bold
Checked = tmpNode.Checked
Expanded = tmpNode.Expanded
ExpandedImage = tmpNode.ExpandedImage
ForeColor = tmpNode.ForeColor
Image = tmpNode.Image
Key = tmpNode.Key
SelectedImage = tmpNode.SelectedImage
Sorted = tmpNode.Sorted
Tag = tmpNode.Tag
Text_ = tmpNode.Text
Visible = tmpNode.Visible

If Not tmpNode.Children = 0 Then
    Set ChildNode = New clsNode
    ChildNode.CopyNode tmpNode.Child
    mvarChildNodes.Add ChildNode
    Set cNode = tmpNode.Child
    Do While Not cNode.Next Is Nothing
        Set ChildNode = New clsNode
        ChildNode.CopyNode cNode.Next
        mvarChildNodes.Add ChildNode
        Set cNode = cNode.Next
    Loop
End If
End Sub

Public Sub ResetNode(tvwTree As TreeView, Relative, Optional Rel As TreeRelationshipConstants = tvwNext)
Dim tmpNode As Node
Dim ChildNode As clsNode

Set tmpNode = tvwTree.Nodes.Add(Relative, Rel, Key, Text_)
tmpNode.Backcolor = Backcolor
tmpNode.Bold = Bold
tmpNode.Checked = Checked
tmpNode.Expanded = Expanded
tmpNode.ExpandedImage = ExpandedImage
tmpNode.ForeColor = ForeColor
tmpNode.Image = Image
tmpNode.SelectedImage = SelectedImage
tmpNode.Sorted = Sorted
tmpNode.Tag = Tag

For Each ChildNode In ChildNodes
    ChildNode.ResetNode tvwTree, Key, tvwChild
Next
End Sub


Private Sub Class_Terminate()
    Set mvarChildNodes = Nothing
End Sub



