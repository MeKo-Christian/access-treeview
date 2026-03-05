# Demo Setup

## Prerequisites

- Windows with 64-bit MS Access
- Both DLLs registered (TreeEngine64.dll and TreeViewHost64.dll)

## Demo Table

Create a table `tblTreeNodes` with these columns:

| Column | Type | Description |
|---|---|---|
| NodeID | Text(50) PK | Unique node identifier |
| ParentID | Text(50) | Parent node ID (empty string for root nodes) |
| NodeText | Text(255) | Display text |
| IconKey | Text(50) | Optional icon key |

## Sample Data

```sql
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1', '', 'Company');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.1', '1', 'Engineering');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.2', '1', 'Sales');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.1.1', '1.1', 'Backend Team');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.1.2', '1.1', 'Frontend Team');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.2.1', '1.2', 'DACH Region');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.2.2', '1.2', 'International');
```

## Form Setup

1. Create a new form `frmTreeDemo`
2. Add an ActiveX control → select "MeKo.TreeViewHost" → name it `ctlTreeView`
3. Add a text box `txtSelectedNode` (displays selected node ID)
4. Add a text box `txtSearch` (search input)
5. Add buttons `cmdSearch` and `cmdRefresh`
6. Import `Form_frmTreeDemo.cls` as the form's VBA module
7. Import `modTreeCompat.bas` into the database
