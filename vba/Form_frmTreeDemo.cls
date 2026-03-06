VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTreeDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim WithEvents tvHost As AccessTreeView.TreeViewHostControl

Private Sub Form_Load()
    Set tvHost = Me.ctlTreeView.Object

    Dim eng As Object
    Set eng = CreateObject("Access.TreeEngine")

    ' Point to demo table in this database
    Dim connStr As String
    connStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & CurrentDb.Name & _
              ";Table=tblTreeNodes;IdCol=NodeID;ParentCol=ParentID;CaptionCol=NodeText"

    eng.Initialize connStr
    tvHost.Initialize eng
End Sub

Private Sub tvHost_NodeClick(ByVal nodeId As String)
    Me.txtSelectedNode = nodeId
End Sub

Private Sub tvHost_NodeDoubleClick(ByVal nodeId As String)
    MsgBox "Double-clicked node: " & nodeId
End Sub

Private Sub tvHost_AfterSelect(ByVal nodeId As String)
    Me.txtSelectedNode = nodeId
End Sub

Private Sub tvHost_OnError(ByVal message As String)
    MsgBox "TreeView Error: " & message, vbExclamation
End Sub

Private Sub cmdSearch_Click()
    If Not tvHost.FindAndSelect(Me.txtSearch) Then
        MsgBox "Not found: " & Me.txtSearch
    End If
End Sub

Private Sub cmdRefresh_Click()
    tvHost.Reload
End Sub
