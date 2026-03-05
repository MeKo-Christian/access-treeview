# TreeView COM Contract

## ITreeEngine (ProgId: MeKo.TreeEngine)

### Methods
| Method | Signature | Description |
|---|---|---|
| Initialize | `(connectionString As String, Optional context As Variant)` | Connect to data source |
| GetRootNodes | `() As ITreeNodeCollection` | Return top-level nodes |
| GetChildren | `(nodeId As String) As ITreeNodeCollection` | Return children of a node |
| HasChildren | `(nodeId As String) As Boolean` | Check if node has children |
| Find | `(text As String, Optional maxResults As Long = 100) As ITreeNodeCollection` | Search nodes by text |
| GetNode | `(nodeId As String) As ITreeNode` | Get a single node by ID |
| Invalidate | `(nodeId As String)` | Clear cached children for a node |
| Reload | `()` | Clear all caches, reload root |

## ITreeNode

| Property | Type | Access | Description |
|---|---|---|---|
| Id | String | get | Unique node identifier |
| ParentId | String | get | Parent node ID (empty for root) |
| Caption | String | get/set | Display text |
| IconKey | String | get/set | Image key for icon |
| Tag | Variant | get/set | User-defined data |

## ITreeNodeCollection

| Member | Description |
|---|---|
| Count | Number of nodes |
| Item(index) | Get node by 1-based index (VBA convention) |
| _NewEnum | For Each support in VBA |

## ITreeViewHost (ProgId: MeKo.TreeViewHost)

### Properties
| Property | Type | Access | Description |
|---|---|---|---|
| Engine | Object | get/set | The bound ITreeEngine instance |
| SelectedNodeId | String | get | Currently selected node ID |
| CheckBoxes | Boolean | get/set | Enable/disable checkboxes |

### Methods
| Method | Signature | Description |
|---|---|---|
| Initialize | `(engine As Object)` | Bind engine to control |
| Reload | `()` | Reload tree from engine |
| ExpandNode | `(nodeId As String)` | Expand a specific node |
| CollapseNode | `(nodeId As String)` | Collapse a specific node |
| SelectNode | `(nodeId As String)` | Select a specific node |
| FindAndSelect | `(text As String) As Boolean` | Search and select first match |

### Events (source interface for WithEvents)
| Event | DispId | Signature | Description |
|---|---|---|---|
| NodeClick | 1 | `(nodeId As String)` | Single click on node |
| NodeDoubleClick | 2 | `(nodeId As String)` | Double click on node |
| BeforeExpand | 3 | `(nodeId As String, ByRef Cancel As Boolean)` | Before expanding (cancelable) |
| AfterExpand | 4 | `(nodeId As String)` | After expanding |
| AfterCollapse | 5 | `(nodeId As String)` | After collapsing |
| AfterSelect | 6 | `(nodeId As String)` | After selection changes |
| OnError | 7 | `(message As String)` | Error notification |

## VBA Usage Example

```vba
Dim WithEvents tvHost As TreeViewHost64.TreeViewHostControl

Private Sub Form_Load()
    Dim eng As Object
    Set eng = CreateObject("MeKo.TreeEngine")
    eng.Initialize "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & CurrentDb.Name

    Set tvHost = Me.ctlTreeView.Object
    tvHost.Initialize eng
End Sub

Private Sub tvHost_AfterSelect(ByVal nodeId As String)
    Debug.Print "Selected: " & nodeId
End Sub

Private Sub tvHost_NodeDoubleClick(ByVal nodeId As String)
    MsgBox "Double-clicked: " & nodeId
End Sub
```
