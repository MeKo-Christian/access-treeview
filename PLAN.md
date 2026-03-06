# Access TreeView — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.
> **Detailed plan:** [docs/plans/2026-03-05-access-treeview.md](docs/plans/2026-03-05-access-treeview.md)

**Goal:** Build a 64-bit TreeView ActiveX control for MS Access, backed by a COM engine DLL, replacing the legacy MSCOMCTL TreeView.

**Architecture:** Two-component design — AccessTreeEngine (data/logic COM DLL) + AccessTreeView (visual ActiveX control wrapping WinForms TreeView). Both registered as COM servers for VBA scripting via `WithEvents`.

**Tech Stack:** C# (.NET 10.0 for dev, .NET Framework 4.8 for production), WinForms, COM Interop, WiX Toolset (installer)

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

- [x] Confirm language choice
- [x] Write ADR `docs/decisions/0001-language-csharp.md`
- [x] Commit

### Task 0.2: Create Solution Structure

- [x] Create `MeKoTreeView.slnx` (SDK used .slnx format)
- [x] Create `src/AccessTreeEngine/AccessTreeEngine.csproj` (net10.0, x64)
- [x] Create `src/AccessTreeView/AccessTreeView.csproj` (net10.0-windows, x64, WinForms)
- [x] Create `tests/AccessTreeEngine.Tests/AccessTreeEngine.Tests.csproj` (NUnit)
- [x] Add projects to solution
- [x] Add .gitignore (standard .NET)
- [x] Verify `dotnet build MeKoTreeView.slnx` succeeds
- [x] Commit

### Task 0.3: Define the COM Contract

- [x] Write `docs/com-contract.md` with full API surface:
  - ITreeEngine: Initialize, GetRootNodes, GetChildren, HasChildren, Find, GetNode, Invalidate, Reload
  - ITreeNode: Id, ParentId, Caption, IconKey, Tag
  - ITreeNodeCollection: Count, Item(index), _NewEnum
  - ITreeViewHost: Engine, SelectedNodeId, Initialize, Reload, ExpandNode, CollapseNode, SelectNode, FindAndSelect
  - Events: NodeClick, NodeDoubleClick, BeforeExpand, AfterExpand, AfterCollapse, AfterSelect, OnError
- [ ] Cross-check against existing VBA TreeView usage in IPOffice (if applicable)
- [x] Commit

---

## Phase 1 — AccessTreeEngine COM DLL

### Task 1.1: Define COM Interfaces

- [x] Write failing test: `ITreeNode` has required properties (Id, ParentId, Caption, IconKey, Tag)
- [x] Write failing test: `ITreeEngine` has required methods
- [x] Run tests — verify FAIL
- [x] Implement `ITreeNode.cs` with `[ComVisible]`, `[Guid]`, `[InterfaceType(InterfaceIsIDispatch)]`
- [x] Implement `ITreeNodeCollection.cs` with `[DispId(-4)]` for `_NewEnum`
- [x] Implement `ITreeEngine.cs`
- [x] Run tests — verify PASS
- [x] Commit

### Task 1.2: Implement TreeNode

- [x] Write failing test: TreeNode stores properties, defaults, Tag stores arbitrary object
- [x] Run test — verify FAIL
- [x] Implement `TreeNode.cs` with `[ComVisible]`, `[ClassInterface(None)]`
- [x] Run test — verify PASS
- [x] Commit

### Task 1.3: Implement TreeNodeCollection

- [x] Write failing test: empty collection count=0, 1-based indexing, ForEach enumeration
- [x] Run test — verify FAIL
- [x] Implement `TreeNodeCollection.cs` (1-based indexing for VBA)
- [x] Run test — verify PASS
- [x] Commit

### Task 1.4: Implement TreeEngine Core (In-Memory Provider)

- [x] Write failing tests: GetRootNodes, GetChildren, HasChildren, GetNode, Find, Find with maxResults
- [x] Run tests — verify FAIL
- [x] Create `ITreeDataProvider.cs` interface
- [x] Implement `InMemoryProvider.cs`
- [x] Implement `TreeEngine.cs` with `[ProgId("Access.TreeEngine")]`
- [x] Run tests — verify PASS
- [x] Commit

### Task 1.5: Implement Caching Layer

- [x] Write failing test: GetChildren caches results, Invalidate clears cache
- [x] Run test — verify FAIL
- [x] Implement `CachingProviderDecorator.cs` (wraps any `ITreeDataProvider`)
- [x] Run test — verify PASS
- [x] Commit

### Task 1.6: Implement OleDb Data Provider

- [x] Write unit test: DbProvider stores connection string
- [x] Run test — verify FAIL
- [x] Implement `DbProvider.cs` (parameterized queries, configurable column names, DbProviderFactory pattern)
- [x] Run test — verify PASS
- [x] Wire `TreeEngine.Initialize()` to create DbProvider + CachingProviderDecorator
- [x] Commit

### Task 1.7: COM Registration Smoke Test

- [ ] Build in Release mode *(requires Windows)*
- [ ] Register with `regasm /codebase /tlb` *(requires Windows)*
- [ ] Test from VBA: `CreateObject("Access.TreeEngine")` returns object *(requires Windows)*
- [ ] Commit

---

## Phase 2 — AccessTreeView Visual ActiveX Control

### Task 2.1: Define Host Interfaces

- [x] Create `ITreeViewHost.cs` (COM-visible interface with all methods/properties)
- [x] Create `ITreeViewHostEvents.cs` (source interface with DispIds for `WithEvents`)
- [x] Commit

### Task 2.2: Create WinForms UserControl Shell

- [x] Implement `TreeViewHostControl.cs`:
  - `[ComVisible]`, `[ProgId("MeKo.TreeViewHost")]`, `[ComSourceInterfaces]`
  - Embedded `System.Windows.Forms.TreeView` (Dock=Fill)
  - Event delegates matching ITreeViewHostEvents
  - Stub methods for ITreeViewHost
- [ ] Build — verify compiles *(requires Windows — WinForms)*
- [x] Commit

### Task 2.3: Implement Tree Loading (Root + Lazy Children)

- [x] Implement `Reload()` — clear tree, load root nodes from engine
- [x] Implement `CreateVisualNode()` — add dummy "Loading..." child when HasChildren=true
- [x] Implement `TreeView_BeforeExpand` — replace dummy with real children from engine
- [x] Use `BeginUpdate/EndUpdate` around bulk node operations
- [ ] Build — verify compiles *(requires Windows)*
- [x] Commit

### Task 2.4: Wire Remaining Events

- [x] Implement `TreeView_AfterCollapse` → raise `AfterCollapse`
- [x] Implement `TreeView_AfterSelect` → raise `AfterSelect`
- [x] Implement `TreeView_NodeMouseClick` → raise `NodeClick`
- [x] Implement `TreeView_NodeMouseDoubleClick` → raise `NodeDoubleClick`
- [ ] Build — verify compiles *(requires Windows)*
- [x] Commit

### Task 2.5: Implement FindAndSelect

- [x] Implement `FindAndSelect(text)` — call engine.Find, expand parent chain, select node
- [x] Implement `ExpandParentChain(nodeId)` — walk up via engine.GetNode, expand from root down
- [ ] Build — verify compiles *(requires Windows)*
- [x] Commit

### Task 2.6: Optional Features (Checkboxes, ImageList)

- [x] Add `CheckBoxes` property to ITreeViewHost and control
- [ ] Add `SetImageList()` support *(deferred — needs design for COM image list transfer)*
- [ ] Build — verify compiles *(requires Windows)*
- [x] Commit

### Task 2.7: COM Registration & Access Form Test

- [ ] Build both DLLs in Release mode *(requires Windows)*
- [ ] Register both with `regasm /codebase /tlb` *(requires Windows)*
- [ ] In Access: insert ActiveX control on form *(requires Windows)*
- [ ] Write VBA: `WithEvents tvHost`, Initialize in Form_Load, handle AfterSelect *(requires Windows)*
- [ ] Verify events fire in VBA Immediate window *(requires Windows)*
- [ ] Commit

---

## Phase 3 — Access Integration

### Task 3.1: Create VBA Wrapper Module

- [x] Write `vba/modTreeCompat.bas`:
  - `Tree_Init(ctl, connectionString, tableName, idCol, parentCol, captionCol)`
  - `Tree_SelectByKey(ctl, nodeId)`
  - `Tree_Refresh(ctl)`
  - `Tree_SelectedKey(ctl) As String`
  - `Tree_FindAndSelect(ctl, searchText) As Boolean`
- [x] Commit

### Task 3.2: Create Demo Access Form

- [x] Write `vba/Form_frmTreeDemo.cls` with WithEvents, Form_Load, event handlers
- [x] Define demo table `tblTreeNodes` schema (NodeID, ParentID, NodeText, IconKey)
- [x] Create sample hierarchical data (7+ rows, 3 levels deep)
- [ ] Create `demo/DemoTreeView.accdb` with table + form *(manual step on Windows)*
- [ ] Verify demo works end-to-end *(requires Windows)*
- [x] Commit

---

## Phase 4 — Installer & Deployment

### Task 4.1: Create WiX Installer Project

- [ ] Install WiX Toolset (`dotnet tool install --global wix`) *(on Windows)*
- [x] Write `installer/Product.wxs` with Package, MajorUpgrade, Feature, ComponentGroups
- [x] Write `installer/README.md` with build instructions
- [ ] Build installer *(requires Windows)*
- [x] Commit

### Task 4.2: Add Registration Custom Actions

- [x] Add `regasm /codebase` custom action on install for both DLLs
- [x] Add `regasm /unregister` custom action on uninstall
- [ ] Test install on clean Windows VM — `CreateObject("Access.TreeEngine")` works *(requires Windows)*
- [ ] Test uninstall — CreateObject fails as expected *(requires Windows)*
- [x] Commit

---

## Phase 5 — Testing

### Task 5.1: Unit Test Edge Cases

- [x] Empty tree (no root nodes)
- [x] Node with no children
- [x] Find with no results
- [x] Find with special characters
- [x] GetNode with null/empty ID
- [x] Very long caption text
- [x] Caching invalidation of nonexistent keys
- [x] Run all tests — 57 pass, 0 fail
- [x] Commit

### Task 5.2: Integration Tests (Windows Only)

- [x] OleDbProvider loads root nodes from real .accdb *(test written, skipped on Linux)*
- [x] OleDbProvider loads children *(test written, skipped on Linux)*
- [x] OleDbProvider handles empty table *(test written, skipped on Linux)*
- [x] Full round-trip: engine → provider → cache → retrieve *(runs on Linux with InMemoryProvider)*
- [ ] Run integration tests on Windows — all pass *(requires Windows)*
- [x] Commit

### Task 5.3: Access Manual Test Checklist

- [x] Write `docs/test-checklist.md`:
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
- [x] Commit
