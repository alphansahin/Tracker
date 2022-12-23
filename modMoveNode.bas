Attribute VB_Name = "modMoveNode"
Option Explicit

Function MoveNode(tvwTree As TreeView, tvwNode As Node, Optional MoveUp As Boolean = True) As Boolean

Dim tmpNode As New clsNode
Dim tvwNodeKey
Dim MoveKey
Dim SelectedKey
Dim Rel As TreeRelationshipConstants
On Error GoTo MoveNodeError

tvwNodeKey = tvwNode.Key

tmpNode.CopyNode tvwNode
If MoveUp Then
    If Not tvwNode.Previous Is Nothing Then
        MoveKey = tvwNode.Previous.Key
        Rel = tvwPrevious
    Else
        MoveNode = False
        Exit Function
    End If
Else
    If Not tvwNode.Next Is Nothing Then
        MoveKey = tvwNode.Next.Key
        Rel = tvwNext
    Else
        MoveNode = False
        Exit Function
    End If
End If

tvwTree.Nodes.Remove tvwNodeKey
tmpNode.ResetNode tvwTree, MoveKey, Rel
tvwTree.SelectedItem = tvwTree.Nodes(tvwNodeKey)
tvwTree.SelectedItem.EnsureVisible

MoveNode = True
On Error GoTo 0
Exit Function


MoveNodeError:
    MoveNode = False
    On Error GoTo 0
End Function


