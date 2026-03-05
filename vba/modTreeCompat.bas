Attribute VB_Name = "modTreeCompat"
Option Compare Database
Option Explicit

' modTreeCompat - Drop-in helpers for MeKo TreeView
' Usage:
'   Tree_Init Me.ctlTree, "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=..."
'   Tree_SelectByKey Me.ctlTree, "42"
'   Tree_Refresh Me.ctlTree

Public Sub Tree_Init(ctl As Object, connectionString As String, _
                     Optional tableName As String = "tblTreeNodes", _
                     Optional idCol As String = "NodeID", _
                     Optional parentCol As String = "ParentID", _
                     Optional captionCol As String = "NodeText")
    Dim eng As Object
    Set eng = CreateObject("MeKo.TreeEngine")

    ' Build extended connection string with column mappings
    Dim extConn As String
    extConn = connectionString & _
              ";Table=" & tableName & _
              ";IdCol=" & idCol & _
              ";ParentCol=" & parentCol & _
              ";CaptionCol=" & captionCol

    eng.Initialize extConn
    ctl.Object.Initialize eng
End Sub

Public Sub Tree_SelectByKey(ctl As Object, nodeId As String)
    ctl.Object.SelectNode nodeId
End Sub

Public Sub Tree_Refresh(ctl As Object)
    ctl.Object.Reload
End Sub

Public Function Tree_SelectedKey(ctl As Object) As String
    Tree_SelectedKey = ctl.Object.SelectedNodeId
End Function

Public Function Tree_FindAndSelect(ctl As Object, searchText As String) As Boolean
    Tree_FindAndSelect = ctl.Object.FindAndSelect(searchText)
End Function
