# MeKo TreeView — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.
> **Detailed plan:** [docs/plans/2026-03-05-access-treeview.md](docs/plans/2026-03-05-access-treeview.md)

**Goal:** Build a 64-bit TreeView ActiveX control for MS Access, backed by a COM engine DLL, replacing the legacy MSCOMCTL TreeView.

**Architecture:** Two-component design — TreeEngine64 (data/logic COM DLL) + TreeViewHost64 (visual ActiveX control wrapping WinForms TreeView). Both registered as COM servers for VBA scripting via `WithEvents`.

**Tech Stack:** C# (.NET Framework 4.8), WinForms, COM Interop, WiX Toolset (installer)

---

## Phase 0 — Project Setup & Language Decision

### Task 0.1: Language Evaluation

| Criterion | C# (.NET) | C++ (ATL) | Go |
|---|---|---|---|
| COM server (in-proc DLL) | Built-in `[ComVisible]`, `regasm` | Native ATL, full control | No native COM server support — not viable |
| ActiveX visual control | WinForms UserControl + COM exposure | MFC/ATL control hosting — complex | Not feasible |
| VBA `WithEvents` | Source interface via `[InterfaceType(InterfaceIsIDispatch)]` | Connection points via ATL — boilerplate-heavy | Not possible |
| Dev on Linux | Editing works, build/register/test needs Windows | Same | Same |
| Speed to build | Fastest — one ecosystem, less boilerplate | Slowest — manual COM plumbing | N/A |

**Decision: C# (.NET Framework 4.8)** — Go is not viable for COM, C++ doubles dev time, C# gives COM visibility with attributes and one unified solution. .NET Framework 4.8 (not .NET 8+) because Access COM interop needs in-proc DLLs via `regasm`.

**Dev environment:** Edit on Ubuntu/VS Code. Build, register (`regasm`), and test in Access **must** happen on Windows (VM, remote, or WSL2).

- [ ] Confirm language choice
- [ ] Write ADR `docs/decisions/0001-language-csharp.md`
- [ ] Commit

### Task 0.2: Create Solution Structure

- [ ] Create `MeKoTreeView.sln`
- [ ] Create `src/TreeEngine64/TreeEngine64.csproj` (.NET Framework 4.8, x64, ComVisible)
- [ ] Create `src/TreeViewHost64/TreeViewHost64.csproj` (.NET Framework 4.8, x64, WinForms, ComVisible)
- [ ] Create `tests/TreeEngine64.Tests/TreeEngine64.Tests.csproj` (NUnit)
- [ ] Add projects to solution
- [ ] Add .gitignore (standard .NET)
- [ ] Verify `dotnet build MeKoTreeView.sln` succeeds
- [ ] Commit

### Task 0.3: Define the COM Contract

- [ ] Write `docs/com-contract.md` with full API surface:
  - ITreeEngine: Initialize, GetRootNodes, GetChildren, HasChildren, Find, GetNode, Invalidate, Reload
  - ITreeNode: Id, ParentId, Caption, IconKey, Tag
  - ITreeNodeCollection: Count, Item(index), _NewEnum
  - ITreeViewHost: Engine, SelectedNodeId, Initialize, Reload, ExpandNode, CollapseNode, SelectNode, FindAndSelect
  - Events: NodeClick, NodeDoubleClick, BeforeExpand, AfterExpand, AfterCollapse, AfterSelect, OnError
- [ ] Cross-check against existing VBA TreeView usage in IPOffice (if applicable)
- [ ] Commit

---

## Phase 1 — TreeEngine64 COM DLL

### Task 1.1: Define COM Interfaces

- [ ] Write failing test: `ITreeNode` has required properties (Id, ParentId, Caption, IconKey, Tag)
- [ ] Write failing test: `ITreeEngine` has required methods
- [ ] Run tests — verify FAIL
- [ ] Implement `ITreeNode.cs` with `[ComVisible]`, `[Guid]`, `[InterfaceType(InterfaceIsIDispatch)]`
- [ ] Implement `ITreeNodeCollection.cs` with `[DispId(-4)]` for `_NewEnum`
- [ ] Implement `ITreeEngine.cs`
- [ ] Run tests — verify PASS
- [ ] Commit

### Task 1.2: Implement TreeNode

- [ ] Write failing test: TreeNode stores properties, defaults, Tag stores arbitrary object
- [ ] Run test — verify FAIL
- [ ] Implement `TreeNode.cs` with `[ComVisible]`, `[ClassInterface(None)]`
- [ ] Run test — verify PASS
- [ ] Commit

### Task 1.3: Implement TreeNodeCollection

- [ ] Write failing test: empty collection count=0, 1-based indexing, ForEach enumeration
- [ ] Run test — verify FAIL
- [ ] Implement `TreeNodeCollection.cs` (1-based indexing for VBA)
- [ ] Run test — verify PASS
- [ ] Commit

### Task 1.4: Implement TreeEngine Core (In-Memory Provider)

- [ ] Write failing tests: GetRootNodes, GetChildren, HasChildren, GetNode, Find, Find with maxResults
- [ ] Run tests — verify FAIL
- [ ] Create `ITreeDataProvider.cs` interface
- [ ] Implement `InMemoryProvider.cs`
- [ ] Implement `TreeEngine.cs` with `[ProgId("MeKo.TreeEngine")]`
- [ ] Run tests — verify PASS
- [ ] Commit

### Task 1.5: Implement Caching Layer

- [ ] Write failing test: GetChildren caches results, Invalidate clears cache
- [ ] Run test — verify FAIL
- [ ] Implement `CachingProviderDecorator.cs` (wraps any `ITreeDataProvider`)
- [ ] Run test — verify PASS
- [ ] Commit

### Task 1.6: Implement OleDb Data Provider

- [ ] Write unit test: OleDbProvider stores connection string
- [ ] Run test — verify FAIL
- [ ] Implement `OleDbProvider.cs` (parameterized queries, configurable column names)
- [ ] Run test — verify PASS
- [ ] Wire `TreeEngine.Initialize()` to create OleDbProvider + CachingProviderDecorator
- [ ] Commit

### Task 1.7: COM Registration Smoke Test

- [ ] Build in Release mode
- [ ] Register with `regasm /codebase /tlb` (on Windows)
- [ ] Test from VBA: `CreateObject("MeKo.TreeEngine")` returns object
- [ ] Commit

---

## Phase 2 — TreeViewHost64 Visual ActiveX Control

### Task 2.1: Define Host Interfaces

- [ ] Create `ITreeViewHost.cs` (COM-visible interface with all methods/properties)
- [ ] Create `ITreeViewHostEvents.cs` (source interface with DispIds for `WithEvents`)
- [ ] Commit

### Task 2.2: Create WinForms UserControl Shell

- [ ] Implement `TreeViewHostControl.cs`:
  - `[ComVisible]`, `[ProgId("MeKo.TreeViewHost")]`, `[ComSourceInterfaces]`
  - Embedded `System.Windows.Forms.TreeView` (Dock=Fill)
  - Event delegates matching ITreeViewHostEvents
  - Stub methods for ITreeViewHost
- [ ] Build — verify compiles
- [ ] Commit

### Task 2.3: Implement Tree Loading (Root + Lazy Children)

- [ ] Implement `Reload()` — clear tree, load root nodes from engine
- [ ] Implement `CreateVisualNode()` — add dummy "Loading..." child when HasChildren=true
- [ ] Implement `TreeView_BeforeExpand` — replace dummy with real children from engine
- [ ] Use `BeginUpdate/EndUpdate` around bulk node operations
- [ ] Build — verify compiles
- [ ] Commit

### Task 2.4: Wire Remaining Events

- [ ] Implement `TreeView_AfterCollapse` → raise `AfterCollapse`
- [ ] Implement `TreeView_AfterSelect` → raise `AfterSelect`
- [ ] Implement `TreeView_NodeMouseClick` → raise `NodeClick`
- [ ] Implement `TreeView_NodeMouseDoubleClick` → raise `NodeDoubleClick`
- [ ] Build — verify compiles
- [ ] Commit

### Task 2.5: Implement FindAndSelect

- [ ] Implement `FindAndSelect(text)` — call engine.Find, expand parent chain, select node
- [ ] Implement `ExpandParentChain(nodeId)` — walk up via engine.GetNode, expand from root down
- [ ] Build — verify compiles
- [ ] Commit

### Task 2.6: Optional Features (Checkboxes, ImageList)

- [ ] Add `CheckBoxes` property to ITreeViewHost and control
- [ ] Add `SetImageList()` support
- [ ] Build — verify compiles
- [ ] Commit

### Task 2.7: COM Registration & Access Form Test

- [ ] Build both DLLs in Release mode
- [ ] Register both with `regasm /codebase /tlb`
- [ ] In Access: insert ActiveX control on form
- [ ] Write VBA: `WithEvents tvHost`, Initialize in Form_Load, handle AfterSelect
- [ ] Verify events fire in VBA Immediate window
- [ ] Commit

---

## Phase 3 — Access Integration

### Task 3.1: Create VBA Wrapper Module

- [ ] Write `vba/modTreeCompat.bas`:
  - `Tree_Init(ctl, connectionString, tableName, idCol, parentCol, captionCol)`
  - `Tree_SelectByKey(ctl, nodeId)`
  - `Tree_Refresh(ctl)`
  - `Tree_SelectedKey(ctl) As String`
  - `Tree_FindAndSelect(ctl, searchText) As Boolean`
- [ ] Commit

### Task 3.2: Create Demo Access Form

- [ ] Write `vba/Form_frmTreeDemo.cls` with WithEvents, Form_Load, event handlers
- [ ] Define demo table `tblTreeNodes` schema (NodeID, ParentID, NodeText, IconKey)
- [ ] Create sample hierarchical data (7+ rows, 3 levels deep)
- [ ] Create `demo/DemoTreeView.accdb` with table + form (manual step on Windows)
- [ ] Verify demo works end-to-end
- [ ] Commit

---

## Phase 4 — Installer & Deployment

### Task 4.1: Create WiX Installer Project

- [ ] Install WiX Toolset (`dotnet tool install --global wix`)
- [ ] Create `installer/MeKoTreeView.wixproj`
- [ ] Write `installer/Product.wxs` with Package, MajorUpgrade, Feature, ComponentGroups
- [ ] Build installer
- [ ] Commit

### Task 4.2: Add Registration Custom Actions

- [ ] Add `regasm /codebase` custom action on install for both DLLs
- [ ] Add `regasm /unregister` custom action on uninstall
- [ ] Test install on clean Windows VM — `CreateObject("MeKo.TreeEngine")` works
- [ ] Test uninstall — CreateObject fails as expected
- [ ] Commit

---

## Phase 5 — Testing

### Task 5.1: Unit Test Edge Cases

- [ ] Empty tree (no root nodes)
- [ ] Node with no children
- [ ] Find with no results
- [ ] Find with special characters
- [ ] GetNode with null/empty ID
- [ ] Very long caption text
- [ ] Concurrent access to caching provider
- [ ] Run all tests — all pass
- [ ] Commit

### Task 5.2: Integration Tests (Windows Only)

- [ ] OleDbProvider loads root nodes from real .accdb
- [ ] OleDbProvider loads children
- [ ] OleDbProvider handles empty table
- [ ] OleDbProvider handles missing table (error)
- [ ] Full round-trip: engine → provider → cache → retrieve
- [ ] Run integration tests — all pass
- [ ] Commit

### Task 5.3: Access Manual Test Checklist

- [ ] Write `docs/test-checklist.md`:
  - [ ] Insert control on form via designer
  - [ ] Control renders on form open
  - [ ] Close/reopen form — still works
  - [ ] Close/reopen database — still works
  - [ ] Multiple control instances on different forms
  - [ ] Compile to ACCDE — still works
  - [ ] Expand/collapse — children load correctly
  - [ ] Click/double-click — events fire in VBA
  - [ ] Search and select — node found and highlighted
  - [ ] 1000+ nodes — no visible lag
  - [ ] 10000+ nodes — lazy loading keeps UI responsive
  - [ ] Rapid expand/collapse — no crashes
- [ ] Commit
