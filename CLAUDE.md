# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Is

A 64-bit TreeView ActiveX control for Microsoft Access, replacing the legacy MSCOMCTL TreeView. Two COM-visible C# components:

- **AccessTreeEngine** (`Access.TreeEngine`) — data/logic COM DLL: loads tree nodes from a database, caches children, supports search
- **AccessTreeView** (`MeKo.TreeViewHost`) — visual ActiveX control wrapping WinForms TreeView, lazy-loads children on expand, raises VBA-compatible events via `WithEvents`

## Build Commands

```bash
# Engine only (works on Linux and Windows)
dotnet build src/AccessTreeEngine/

# Run tests (NUnit, 57 unit + 7 integration tests)
dotnet test tests/AccessTreeEngine.Tests/

# Run a single test
dotnet test tests/AccessTreeEngine.Tests/ --filter "FullyQualifiedName~TestMethodName"

# Integration tests only (Windows + Access OLEDB provider required)
dotnet test tests/AccessTreeEngine.Tests/ --filter Category=Integration

# Full solution (Windows only — AccessTreeView requires WinForms)
dotnet build AccessTreeView.slnx
```

## Architecture

### Solution Layout (`AccessTreeView.slnx`)

- `src/AccessTreeEngine/` — .NET 10.0, x64. Namespace: `Access.TreeEngine`
- `src/AccessTreeView/` — .NET 10.0-windows, x64, WinForms. Namespace: `MeKo.TreeViewHost`. Depends on AccessTreeEngine.
- `tests/AccessTreeEngine.Tests/` — NUnit tests for the engine

### Engine Internal Architecture

The engine uses a **provider/decorator pattern**:

1. `ITreeDataProvider` — interface for tree data access (GetRootNodes, GetChildren, HasChildren, Find, GetNode)
2. `InMemoryProvider` — in-memory implementation for tests and simple use cases
3. `DbProvider` — OleDb/ODBC database provider with configurable column mappings
4. `CachingProviderDecorator` — wraps any provider, caches `GetChildren()` results using `ConcurrentDictionary`
5. `TreeEngine` — COM-visible facade (`[ProgId("Access.TreeEngine")]`). `Initialize()` creates DbProvider wrapped in CachingProviderDecorator

### COM Interop Conventions

- All public interfaces use `[ComVisible(true)]`, `[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]`
- Classes use `[ClassInterface(ClassInterfaceType.None)]` and implement their COM interface explicitly
- `TreeNodeCollection` uses **1-based indexing** (VBA convention) and exposes `_NewEnum` with `[DispId(-4)]` for `For Each`
- Events use a source interface (`ITreeViewHostEvents`) with explicit `[DispId]` values for VBA `WithEvents`
- COM registration requires Windows: `regasm /codebase /tlb`

### Key Design Decisions

- **C# over C++/Go** — Go has no COM server support; C++ doubles dev time. C# provides COM visibility via attributes. See `docs/decisions/0001-language-csharp.md`
- **.NET 10.0 for development** (builds/tests on Linux), **.NET Framework 4.8 for production** COM deployment (Access needs in-proc DLLs via regasm)
- **Nullable disabled** in source projects (COM interop compatibility), enabled in test project

## Important Files

- `docs/com-contract.md` — full COM API surface (interfaces, methods, events, DispIds)
- `PLAN.md` — implementation plan with task status (links to detailed plan in `docs/plans/`)
- `installer/Product.wxs` — WiX v4 MSI installer with regasm custom actions
- `vba/modTreeCompat.bas` — VBA compatibility wrapper functions
- `vba/Form_frmTreeDemo.cls` — demo form VBA module

## Platform Notes

- AccessTreeEngine builds and tests run on Linux (57 tests pass, 7 integration tests skipped)
- AccessTreeView requires Windows (WinForms dependency)
- COM registration, installer build, and Access integration testing all require Windows
