# Access TreeView

A 64-bit TreeView ActiveX control for Microsoft Access, replacing the legacy MSCOMCTL TreeView (which has no 64-bit version).

## Overview

Two COM components work together:

- **TreeEngine64** — data and logic (COM in-proc DLL). Loads tree nodes from a database, caches children, supports search. VBA calls it via `CreateObject("MeKo.TreeEngine")`.
- **TreeViewHost64** — visual control (WinForms UserControl exposed as ActiveX). Wraps `System.Windows.Forms.TreeView`, lazy-loads children on expand, and raises VBA-compatible events via `WithEvents`.

## Quick Start (VBA)

```vba
Dim WithEvents tvHost As TreeViewHost64.TreeViewHostControl

Private Sub Form_Load()
    Dim eng As Object
    Set eng = CreateObject("MeKo.TreeEngine")
    eng.Initialize "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & CurrentDb.Name & _
                   ";Table=tblTreeNodes;IdCol=NodeID;ParentCol=ParentID;CaptionCol=NodeText"

    Set tvHost = Me.ctlTreeView.Object
    tvHost.Initialize eng
End Sub

Private Sub tvHost_AfterSelect(ByVal nodeId As String)
    Debug.Print "Selected: " & nodeId
End Sub
```

Or use the compatibility wrapper:

```vba
Private Sub Form_Load()
    Tree_Init Me.ctlTreeView, "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & CurrentDb.Name
End Sub
```

## Project Structure

```
src/
  TreeEngine64/          Engine COM DLL (data, caching, search)
  TreeViewHost64/        Visual ActiveX control (WinForms)
tests/
  TreeEngine64.Tests/    Unit + integration tests (NUnit)
vba/
  modTreeCompat.bas      Drop-in VBA helper functions
  Form_frmTreeDemo.cls   Demo form VBA module
installer/
  Product.wxs            WiX installer (regasm registration)
docs/
  com-contract.md        Full API surface definition
  test-checklist.md      Manual Access test checklist
  decisions/             Architecture Decision Records
  plans/                 Implementation plans
```

## Building

### Prerequisites

- .NET 10 SDK (or .NET 8+)
- Windows required for TreeViewHost64 (WinForms) and COM registration

### Engine (builds on Linux or Windows)

```bash
dotnet build src/TreeEngine64/
dotnet test tests/TreeEngine64.Tests/
```

### Full solution (Windows only)

```bash
dotnet build MeKoTreeView.slnx
```

### COM Registration (Windows, admin)

```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase TreeEngine64.dll /tlb
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase TreeViewHost64.dll /tlb
```

### Installer

See [installer/README.md](installer/README.md) for MSI build instructions.

## API

### TreeEngine (ProgId: `MeKo.TreeEngine`)

| Method | Description |
|---|---|
| `Initialize(connStr)` | Connect to database with column mappings |
| `GetRootNodes()` | Top-level nodes |
| `GetChildren(nodeId)` | Children of a node (cached) |
| `HasChildren(nodeId)` | Check without loading children |
| `Find(text, maxResults)` | Search by caption |
| `GetNode(nodeId)` | Single node by ID |
| `Invalidate(nodeId)` | Clear cache for a node |
| `Reload()` | Clear all caches |

### TreeViewHost (ProgId: `MeKo.TreeViewHost`)

| Method / Property | Description |
|---|---|
| `Initialize(engine)` | Bind an engine and load tree |
| `Reload()` | Refresh from engine |
| `SelectNode(nodeId)` | Select a node |
| `FindAndSelect(text)` | Search and select first match |
| `SelectedNodeId` | Currently selected node ID |
| `CheckBoxes` | Enable/disable checkboxes |

### Events (VBA `WithEvents`)

| Event | Signature |
|---|---|
| `NodeClick` | `(nodeId As String)` |
| `NodeDoubleClick` | `(nodeId As String)` |
| `BeforeExpand` | `(nodeId As String, ByRef Cancel As Boolean)` |
| `AfterExpand` | `(nodeId As String)` |
| `AfterCollapse` | `(nodeId As String)` |
| `AfterSelect` | `(nodeId As String)` |
| `OnError` | `(message As String)` |

Full API details: [docs/com-contract.md](docs/com-contract.md)

## Connection String

The `Initialize` method accepts a standard OleDb/ODBC connection string with additional keys:

| Key | Default | Description |
|---|---|---|
| `Table` | `tblTreeNodes` | Table name |
| `IdCol` | `NodeID` | Node ID column |
| `ParentCol` | `ParentID` | Parent ID column |
| `CaptionCol` | `NodeText` | Display text column |
| `IconCol` | *(none)* | Optional icon key column |
| `DbProvider` | `System.Data.OleDb` | ADO.NET provider factory name |

## Tests

```bash
# Unit tests (Linux or Windows)
dotnet test tests/TreeEngine64.Tests/

# Integration tests only (Windows with Access)
dotnet test tests/TreeEngine64.Tests/ --filter Category=Integration
```

57 unit tests, 7 integration tests (require Windows + Access OLEDB provider).

## License

Proprietary — MeKo-Tech
