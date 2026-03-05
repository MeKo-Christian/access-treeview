# MeKo TreeView — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a 64-bit TreeView ActiveX control for MS Access, backed by a COM engine DLL, replacing the legacy MSCOMCTL TreeView.

**Architecture:** Two-component design — a TreeEngine64 COM DLL handling data/logic, and a TreeViewHost64 visual ActiveX control wrapping WinForms TreeView. Both registered as COM servers so Access can host and script them via VBA.

**Tech Stack:** C# (.NET Framework 4.8), WinForms, COM Interop, WiX Toolset (installer)

---

## Phase 0 — Project Setup & Language Decision

### Task 0.1: Language Evaluation

**Context:** The goal suggests C++ (ATL) or C# (.NET). The developer prefers Go but needs COM interop. Here is the evaluation:

| Criterion | C# (.NET) | C++ (ATL) | Go |
|---|---|---|---|
| COM server (in-proc DLL) | Built-in `[ComVisible]`, `regasm` | Native ATL, full control | No native COM server support; requires CGo + manual IUnknown/IDispatch — extremely painful |
| ActiveX visual control | WinForms UserControl + COM exposure is well-documented | MFC/ATL control hosting — complex | Not feasible — Go cannot produce an ActiveX control |
| VBA `WithEvents` | Source interface via `[InterfaceType(InterfaceIsIDispatch)]` — straightforward | Connection points via ATL — works but boilerplate-heavy | Not possible |
| Development on Linux | **Cannot build COM-registered DLLs on Linux.** Must use Windows (or Windows VM/container) for build + test. `dotnet build` works cross-platform but `regasm` and Access testing require Windows. | Same — needs Windows | Same |
| IDE on Ubuntu | VS Code + C# Dev Kit works well for editing. Build/test must target Windows. | VS Code + CMake works for editing. Build requires MSVC. | N/A — Go is ruled out for COM |
| Speed to build | Fastest — one ecosystem, less boilerplate | Slowest — manual COM plumbing | Not viable |
| Ecosystem match | Both components in C#, shared types, one solution | Mixed or all-C++ | N/A |

**Decision: C# (.NET Framework 4.8)** is the clear winner.

- Go is **not viable** for COM in-proc servers or ActiveX controls
- C++ works but doubles development time with manual COM plumbing
- C# gives COM visibility with attributes, WinForms UserControl hosting, and one unified solution
- .NET Framework 4.8 (not .NET 8+) because Access COM interop requires in-proc DLLs registered via `regasm`, which is best supported on .NET Framework

**Development environment note:** Editing can happen on Ubuntu/VS Code. Building, registering, and testing **must** happen on Windows (physical, VM, or remote). Consider a Windows dev VM or use WSL2 on a Windows machine.

**Step 1: Confirm language choice**

Review this evaluation. If C# is accepted, proceed.

**Step 2: Document the decision**

Create `docs/decisions/0001-language-csharp.md`:

```markdown
# ADR 0001: Use C# (.NET Framework 4.8)

## Status: Accepted

## Context
We need COM in-proc DLLs and an ActiveX visual control for 64-bit Access.
Go lacks COM server support. C++ ATL works but is slow to develop.

## Decision
Use C# with .NET Framework 4.8 for both TreeEngine64 and TreeViewHost64.

## Consequences
- Must build and register on Windows (regasm + Access testing)
- Editing can happen on any OS via VS Code
- Single solution, shared types between engine and control
```

**Step 3: Commit**

```bash
git add docs/
git commit -m "docs: ADR 0001 — choose C# for COM components"
```

---

### Task 0.2: Create Solution Structure

**Files:**
- Create: `MeKoTreeView.sln`
- Create: `src/TreeEngine64/TreeEngine64.csproj`
- Create: `src/TreeViewHost64/TreeViewHost64.csproj`
- Create: `tests/TreeEngine64.Tests/TreeEngine64.Tests.csproj`

**Step 1: Create the solution and projects**

```bash
# Create solution
dotnet new sln -n MeKoTreeView

# Engine project — .NET Framework 4.8 Class Library
mkdir -p src/TreeEngine64
# (create .csproj manually — see below)

# Host control project
mkdir -p src/TreeViewHost64
# (create .csproj manually — see below)

# Test project
mkdir -p tests/TreeEngine64.Tests
dotnet new nunit -n TreeEngine64.Tests -o tests/TreeEngine64.Tests
```

`src/TreeEngine64/TreeEngine64.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <AssemblyName>TreeEngine64</AssemblyName>
    <RootNamespace>MeKo.TreeEngine</RootNamespace>
    <GenerateAssemblyInfo>true</GenerateAssemblyInfo>
    <EnableComHosting>true</EnableComHosting>
    <PlatformTarget>x64</PlatformTarget>
  </PropertyGroup>
</Project>
```

`src/TreeViewHost64/TreeViewHost64.csproj`:
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <AssemblyName>TreeViewHost64</AssemblyName>
    <RootNamespace>MeKo.TreeViewHost</RootNamespace>
    <UseWindowsForms>true</UseWindowsForms>
    <EnableComHosting>true</EnableComHosting>
    <PlatformTarget>x64</PlatformTarget>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\TreeEngine64\TreeEngine64.csproj" />
  </ItemGroup>
</Project>
```

**Step 2: Add projects to solution**

```bash
dotnet sln add src/TreeEngine64/TreeEngine64.csproj
dotnet sln add src/TreeViewHost64/TreeViewHost64.csproj
dotnet sln add tests/TreeEngine64.Tests/TreeEngine64.Tests.csproj
```

**Step 3: Add .gitignore**

Use standard .NET gitignore (bin/, obj/, *.user, etc.)

**Step 4: Verify solution builds**

```bash
dotnet build MeKoTreeView.sln
```

Expected: Build succeeded, 0 errors.

**Step 5: Commit**

```bash
git add .
git commit -m "chore: scaffold solution with engine, host, and test projects"
```

---

### Task 0.3: Define the COM Contract (API Surface)

**Files:**
- Create: `docs/com-contract.md`

**Step 1: Write the COM contract document**

```markdown
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

### ITreeNode
| Property | Type | Description |
|---|---|---|
| Id | String | Unique node identifier |
| ParentId | String | Parent node ID (empty for root) |
| Caption | String | Display text |
| IconKey | String | Image key for icon |
| Tag | Variant | User-defined data |

### ITreeNodeCollection
| Member | Description |
|---|---|
| Count | Number of nodes |
| Item(index) | Get node by index (1-based for VBA) |
| _NewEnum | For Each support |

## ITreeViewHost (ProgId: MeKo.TreeViewHost)

### Properties
| Property | Type | Description |
|---|---|---|
| Engine | Object (ITreeEngine) | The bound engine |
| SelectedNodeId | String (get) | Currently selected node ID |

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
| Event | Signature | Description |
|---|---|---|
| NodeClick | `(nodeId As String)` | Single click on node |
| NodeDoubleClick | `(nodeId As String)` | Double click on node |
| BeforeExpand | `(nodeId As String, ByRef Cancel As Boolean)` | Before expanding (can cancel) |
| AfterExpand | `(nodeId As String)` | After expanding |
| AfterCollapse | `(nodeId As String)` | After collapsing |
| AfterSelect | `(nodeId As String)` | After selection changes |
| OnError | `(message As String)` | Error notification |
```

**Step 2: Review contract against existing VBA usage**

Check the existing IPOffice VBA code for TreeView usage patterns. If existing code uses `Nodes.Add`, `Key`, `Text`, etc., note any gaps in the contract.

**Step 3: Commit**

```bash
git add docs/com-contract.md
git commit -m "docs: define COM contract for engine and host control"
```

---

## Phase 1 — TreeEngine64 COM DLL

### Task 1.1: Define COM Interfaces in C#

**Files:**
- Create: `src/TreeEngine64/ITreeNode.cs`
- Create: `src/TreeEngine64/ITreeNodeCollection.cs`
- Create: `src/TreeEngine64/ITreeEngine.cs`

**Step 1: Write the failing test**

`tests/TreeEngine64.Tests/TreeEngineInterfaceTests.cs`:
```csharp
using NUnit.Framework;
using MeKo.TreeEngine;

[TestFixture]
public class TreeEngineInterfaceTests
{
    [Test]
    public void ITreeNode_Has_Required_Properties()
    {
        var type = typeof(ITreeNode);
        Assert.That(type.GetProperty("Id"), Is.Not.Null);
        Assert.That(type.GetProperty("ParentId"), Is.Not.Null);
        Assert.That(type.GetProperty("Caption"), Is.Not.Null);
        Assert.That(type.GetProperty("IconKey"), Is.Not.Null);
        Assert.That(type.GetProperty("Tag"), Is.Not.Null);
    }

    [Test]
    public void ITreeEngine_Has_Required_Methods()
    {
        var type = typeof(ITreeEngine);
        Assert.That(type.GetMethod("GetRootNodes"), Is.Not.Null);
        Assert.That(type.GetMethod("GetChildren"), Is.Not.Null);
        Assert.That(type.GetMethod("HasChildren"), Is.Not.Null);
        Assert.That(type.GetMethod("Find"), Is.Not.Null);
        Assert.That(type.GetMethod("GetNode"), Is.Not.Null);
    }
}
```

**Step 2: Run test to verify it fails**

```bash
dotnet test tests/TreeEngine64.Tests/ -v n
```
Expected: FAIL — types do not exist yet.

**Step 3: Implement the interfaces**

`src/TreeEngine64/ITreeNode.cs`:
```csharp
using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-1111-1111-1111-000000000001")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITreeNode
    {
        string Id { get; }
        string ParentId { get; }
        string Caption { get; set; }
        string IconKey { get; set; }
        object Tag { get; set; }
    }
}
```

`src/TreeEngine64/ITreeNodeCollection.cs`:
```csharp
using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-1111-1111-1111-000000000002")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITreeNodeCollection
    {
        int Count { get; }
        ITreeNode this[int index] { get; }
        [DispId(-4)] // DISPID_NEWENUM for For Each
        IEnumerator GetEnumerator();
    }
}
```

`src/TreeEngine64/ITreeEngine.cs`:
```csharp
using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-1111-1111-1111-000000000003")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITreeEngine
    {
        void Initialize(string connectionString, object context = null);
        ITreeNodeCollection GetRootNodes();
        ITreeNodeCollection GetChildren(string nodeId);
        bool HasChildren(string nodeId);
        ITreeNodeCollection Find(string text, int maxResults = 100);
        ITreeNode GetNode(string nodeId);
        void Invalidate(string nodeId);
        void Reload();
    }
}
```

**Step 4: Run test to verify it passes**

```bash
dotnet test tests/TreeEngine64.Tests/ -v n
```
Expected: PASS

**Step 5: Commit**

```bash
git add src/TreeEngine64/ tests/TreeEngine64.Tests/
git commit -m "feat(engine): define COM interfaces ITreeNode, ITreeNodeCollection, ITreeEngine"
```

---

### Task 1.2: Implement TreeNode

**Files:**
- Create: `src/TreeEngine64/TreeNode.cs`
- Create: `tests/TreeEngine64.Tests/TreeNodeTests.cs`

**Step 1: Write the failing test**

```csharp
using NUnit.Framework;
using MeKo.TreeEngine;

[TestFixture]
public class TreeNodeTests
{
    [Test]
    public void TreeNode_Stores_Properties()
    {
        var node = new TreeNode("42", "10", "My Node");
        Assert.That(node.Id, Is.EqualTo("42"));
        Assert.That(node.ParentId, Is.EqualTo("10"));
        Assert.That(node.Caption, Is.EqualTo("My Node"));
    }

    [Test]
    public void TreeNode_IconKey_Defaults_To_Empty()
    {
        var node = new TreeNode("1", "", "Root");
        Assert.That(node.IconKey, Is.EqualTo(""));
    }

    [Test]
    public void TreeNode_Tag_Can_Store_Arbitrary_Object()
    {
        var node = new TreeNode("1", "", "Root");
        node.Tag = "some data";
        Assert.That(node.Tag, Is.EqualTo("some data"));
    }
}
```

**Step 2: Run test — expect FAIL**

**Step 3: Implement TreeNode**

```csharp
using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-2222-2222-2222-000000000001")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TreeNode : ITreeNode
    {
        public TreeNode(string id, string parentId, string caption)
        {
            Id = id;
            ParentId = parentId ?? "";
            Caption = caption ?? "";
            IconKey = "";
        }

        public string Id { get; }
        public string ParentId { get; }
        public string Caption { get; set; }
        public string IconKey { get; set; }
        public object Tag { get; set; }
    }
}
```

**Step 4: Run test — expect PASS**

**Step 5: Commit**

```bash
git commit -am "feat(engine): implement TreeNode"
```

---

### Task 1.3: Implement TreeNodeCollection

**Files:**
- Create: `src/TreeEngine64/TreeNodeCollection.cs`
- Create: `tests/TreeEngine64.Tests/TreeNodeCollectionTests.cs`

**Step 1: Write the failing test**

```csharp
using NUnit.Framework;
using MeKo.TreeEngine;
using System.Collections.Generic;

[TestFixture]
public class TreeNodeCollectionTests
{
    [Test]
    public void Empty_Collection_Has_Count_Zero()
    {
        var coll = new TreeNodeCollection(new List<TreeNode>());
        Assert.That(coll.Count, Is.EqualTo(0));
    }

    [Test]
    public void Item_Returns_Node_By_OneBasedIndex()
    {
        var nodes = new List<TreeNode>
        {
            new TreeNode("1", "", "First"),
            new TreeNode("2", "", "Second")
        };
        var coll = new TreeNodeCollection(nodes);

        Assert.That(coll[1].Caption, Is.EqualTo("First"));
        Assert.That(coll[2].Caption, Is.EqualTo("Second"));
    }

    [Test]
    public void Supports_ForEach_Enumeration()
    {
        var nodes = new List<TreeNode>
        {
            new TreeNode("1", "", "A"),
            new TreeNode("2", "", "B")
        };
        var coll = new TreeNodeCollection(nodes);
        var captions = new List<string>();

        foreach (ITreeNode node in coll)
            captions.Add(node.Caption);

        Assert.That(captions, Is.EqualTo(new[] { "A", "B" }));
    }
}
```

**Step 2: Run test — expect FAIL**

**Step 3: Implement TreeNodeCollection**

```csharp
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-2222-2222-2222-000000000002")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TreeNodeCollection : ITreeNodeCollection
    {
        private readonly List<TreeNode> _nodes;

        public TreeNodeCollection(List<TreeNode> nodes)
        {
            _nodes = nodes ?? new List<TreeNode>();
        }

        public int Count => _nodes.Count;

        public ITreeNode this[int index] => _nodes[index - 1]; // 1-based for VBA

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _nodes.GetEnumerator();
    }
}
```

**Step 4: Run test — expect PASS**

**Step 5: Commit**

```bash
git commit -am "feat(engine): implement TreeNodeCollection with 1-based indexing"
```

---

### Task 1.4: Implement TreeEngine Core (In-Memory Provider)

**Files:**
- Create: `src/TreeEngine64/ITreeDataProvider.cs`
- Create: `src/TreeEngine64/TreeEngine.cs`
- Create: `tests/TreeEngine64.Tests/TreeEngineTests.cs`

**Step 1: Write the failing test**

```csharp
using NUnit.Framework;
using MeKo.TreeEngine;
using System.Collections.Generic;

[TestFixture]
public class TreeEngineTests
{
    private TreeEngine _engine;

    [SetUp]
    public void Setup()
    {
        _engine = new TreeEngine();
        _engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Root A"),
            new TreeNode("2", "", "Root B"),
            new TreeNode("1.1", "1", "Child A1"),
            new TreeNode("1.2", "1", "Child A2"),
            new TreeNode("2.1", "2", "Child B1"),
        }));
    }

    [Test]
    public void GetRootNodes_Returns_Top_Level_Nodes()
    {
        var roots = _engine.GetRootNodes();
        Assert.That(roots.Count, Is.EqualTo(2));
        Assert.That(roots[1].Caption, Is.EqualTo("Root A"));
        Assert.That(roots[2].Caption, Is.EqualTo("Root B"));
    }

    [Test]
    public void GetChildren_Returns_Children_Of_Node()
    {
        var children = _engine.GetChildren("1");
        Assert.That(children.Count, Is.EqualTo(2));
        Assert.That(children[1].Caption, Is.EqualTo("Child A1"));
    }

    [Test]
    public void HasChildren_Returns_True_For_Parent_Node()
    {
        Assert.That(_engine.HasChildren("1"), Is.True);
    }

    [Test]
    public void HasChildren_Returns_False_For_Leaf_Node()
    {
        Assert.That(_engine.HasChildren("1.1"), Is.False);
    }

    [Test]
    public void GetNode_Returns_Specific_Node()
    {
        var node = _engine.GetNode("1.2");
        Assert.That(node.Caption, Is.EqualTo("Child A2"));
    }

    [Test]
    public void GetNode_Returns_Null_For_Unknown_Id()
    {
        var node = _engine.GetNode("999");
        Assert.That(node, Is.Null);
    }

    [Test]
    public void Find_Returns_Matching_Nodes()
    {
        var results = _engine.Find("Child A");
        Assert.That(results.Count, Is.EqualTo(2));
    }

    [Test]
    public void Find_Respects_MaxResults()
    {
        var results = _engine.Find("Child", 1);
        Assert.That(results.Count, Is.EqualTo(1));
    }
}
```

**Step 2: Run test — expect FAIL**

**Step 3: Implement the data provider interface and in-memory provider**

`src/TreeEngine64/ITreeDataProvider.cs`:
```csharp
using System.Collections.Generic;

namespace MeKo.TreeEngine
{
    public interface ITreeDataProvider
    {
        List<TreeNode> GetRootNodes();
        List<TreeNode> GetChildren(string parentId);
        bool HasChildren(string nodeId);
        TreeNode GetNode(string nodeId);
        List<TreeNode> Find(string text, int maxResults);
    }
}
```

`src/TreeEngine64/InMemoryProvider.cs`:
```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace MeKo.TreeEngine
{
    public class InMemoryProvider : ITreeDataProvider
    {
        private readonly List<TreeNode> _allNodes;

        public InMemoryProvider(List<TreeNode> nodes)
        {
            _allNodes = nodes ?? new List<TreeNode>();
        }

        public List<TreeNode> GetRootNodes()
            => _allNodes.Where(n => string.IsNullOrEmpty(n.ParentId)).ToList();

        public List<TreeNode> GetChildren(string parentId)
            => _allNodes.Where(n => n.ParentId == parentId).ToList();

        public bool HasChildren(string nodeId)
            => _allNodes.Any(n => n.ParentId == nodeId);

        public TreeNode GetNode(string nodeId)
            => _allNodes.FirstOrDefault(n => n.Id == nodeId);

        public List<TreeNode> Find(string text, int maxResults)
            => _allNodes
                .Where(n => n.Caption.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                .Take(maxResults)
                .ToList();
    }
}
```

**Step 4: Implement TreeEngine**

`src/TreeEngine64/TreeEngine.cs`:
```csharp
using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-3333-3333-3333-000000000001")]
    [ProgId("MeKo.TreeEngine")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TreeEngine : ITreeEngine
    {
        private ITreeDataProvider _provider;

        public void SetProvider(ITreeDataProvider provider)
        {
            _provider = provider;
        }

        public void Initialize(string connectionString, object context = null)
        {
            // Phase 1.5 will add OleDb/ODBC provider
            // For now, Initialize is a no-op if provider is already set
        }

        public ITreeNodeCollection GetRootNodes()
        {
            EnsureProvider();
            return new TreeNodeCollection(_provider.GetRootNodes());
        }

        public ITreeNodeCollection GetChildren(string nodeId)
        {
            EnsureProvider();
            return new TreeNodeCollection(_provider.GetChildren(nodeId));
        }

        public bool HasChildren(string nodeId)
        {
            EnsureProvider();
            return _provider.HasChildren(nodeId);
        }

        public ITreeNodeCollection Find(string text, int maxResults = 100)
        {
            EnsureProvider();
            return new TreeNodeCollection(_provider.Find(text, maxResults));
        }

        public ITreeNode GetNode(string nodeId)
        {
            EnsureProvider();
            return _provider.GetNode(nodeId);
        }

        public void Invalidate(string nodeId)
        {
            // Will be meaningful with caching provider
        }

        public void Reload()
        {
            // Will be meaningful with DB provider
        }

        private void EnsureProvider()
        {
            if (_provider == null)
                throw new InvalidOperationException(
                    "TreeEngine not initialized. Call SetProvider() or Initialize() first.");
        }
    }
}
```

**Step 5: Run test — expect PASS**

**Step 6: Commit**

```bash
git commit -am "feat(engine): implement TreeEngine with InMemoryProvider"
```

---

### Task 1.5: Implement Caching Layer

**Files:**
- Create: `src/TreeEngine64/CachingProviderDecorator.cs`
- Create: `tests/TreeEngine64.Tests/CachingProviderTests.cs`

**Step 1: Write the failing test**

```csharp
using NUnit.Framework;
using MeKo.TreeEngine;
using System.Collections.Generic;

[TestFixture]
public class CachingProviderTests
{
    [Test]
    public void GetChildren_Caches_Results()
    {
        var callCount = 0;
        var inner = new CountingProvider(() => callCount++);
        var caching = new CachingProviderDecorator(inner);

        caching.GetChildren("1");
        caching.GetChildren("1");

        Assert.That(callCount, Is.EqualTo(1), "Second call should use cache");
    }

    [Test]
    public void Invalidate_Clears_Cache_For_Node()
    {
        var callCount = 0;
        var inner = new CountingProvider(() => callCount++);
        var caching = new CachingProviderDecorator(inner);

        caching.GetChildren("1");
        caching.Invalidate("1");
        caching.GetChildren("1");

        Assert.That(callCount, Is.EqualTo(2), "Should re-fetch after invalidation");
    }

    // Helper: a provider that counts GetChildren calls
    private class CountingProvider : InMemoryProvider
    {
        private readonly System.Action _onGetChildren;

        public CountingProvider(System.Action onGetChildren)
            : base(new List<TreeNode>
            {
                new TreeNode("1", "", "Root"),
                new TreeNode("1.1", "1", "Child")
            })
        {
            _onGetChildren = onGetChildren;
        }

        public new List<TreeNode> GetChildren(string parentId)
        {
            _onGetChildren();
            return base.GetChildren(parentId);
        }
    }
}
```

**Step 2: Run test — expect FAIL**

**Step 3: Implement CachingProviderDecorator**

```csharp
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace MeKo.TreeEngine
{
    public class CachingProviderDecorator : ITreeDataProvider
    {
        private readonly ITreeDataProvider _inner;
        private readonly ConcurrentDictionary<string, List<TreeNode>> _childrenCache = new();

        public CachingProviderDecorator(ITreeDataProvider inner)
        {
            _inner = inner;
        }

        public List<TreeNode> GetRootNodes() => _inner.GetRootNodes();

        public List<TreeNode> GetChildren(string parentId)
        {
            return _childrenCache.GetOrAdd(parentId, id => _inner.GetChildren(id));
        }

        public bool HasChildren(string nodeId) => _inner.HasChildren(nodeId);
        public TreeNode GetNode(string nodeId) => _inner.GetNode(nodeId);
        public List<TreeNode> Find(string text, int maxResults) => _inner.Find(text, maxResults);

        public void Invalidate(string nodeId)
        {
            _childrenCache.TryRemove(nodeId, out _);
        }

        public void InvalidateAll()
        {
            _childrenCache.Clear();
        }
    }
}
```

**Step 4: Run test — expect PASS**

**Step 5: Commit**

```bash
git commit -am "feat(engine): add CachingProviderDecorator for lazy-load caching"
```

---

### Task 1.6: Implement OleDb Data Provider

**Files:**
- Create: `src/TreeEngine64/OleDbProvider.cs`
- Create: `tests/TreeEngine64.Tests/OleDbProviderTests.cs`

**Step 1: Write the failing test**

```csharp
using NUnit.Framework;
using MeKo.TreeEngine;

[TestFixture]
public class OleDbProviderTests
{
    [Test]
    public void Constructor_Stores_ConnectionString()
    {
        var provider = new OleDbProvider(
            connectionString: "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb",
            tableName: "tblTree",
            idColumn: "NodeID",
            parentIdColumn: "ParentID",
            captionColumn: "NodeText"
        );

        Assert.That(provider.ConnectionString, Is.Not.Null);
    }

    // Integration tests (require actual Access DB) go in a separate fixture
    // marked [Category("Integration")] and skipped in CI
}
```

**Step 2: Run test — expect FAIL**

**Step 3: Implement OleDbProvider**

```csharp
using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace MeKo.TreeEngine
{
    public class OleDbProvider : ITreeDataProvider
    {
        public string ConnectionString { get; }
        private readonly string _tableName;
        private readonly string _idCol;
        private readonly string _parentIdCol;
        private readonly string _captionCol;
        private readonly string _iconKeyCol;

        public OleDbProvider(string connectionString, string tableName,
            string idColumn, string parentIdColumn, string captionColumn,
            string iconKeyColumn = null)
        {
            ConnectionString = connectionString;
            _tableName = tableName;
            _idCol = idColumn;
            _parentIdCol = parentIdColumn;
            _captionCol = captionColumn;
            _iconKeyCol = iconKeyColumn;
        }

        public List<TreeNode> GetRootNodes()
        {
            return QueryNodes($"SELECT * FROM [{_tableName}] WHERE [{_parentIdCol}] IS NULL OR [{_parentIdCol}] = ''");
        }

        public List<TreeNode> GetChildren(string parentId)
        {
            return QueryNodes($"SELECT * FROM [{_tableName}] WHERE [{_parentIdCol}] = @parentId",
                new OleDbParameter("@parentId", parentId));
        }

        public bool HasChildren(string nodeId)
        {
            using var conn = new OleDbConnection(ConnectionString);
            conn.Open();
            using var cmd = new OleDbCommand(
                $"SELECT COUNT(*) FROM [{_tableName}] WHERE [{_parentIdCol}] = @nodeId", conn);
            cmd.Parameters.AddWithValue("@nodeId", nodeId);
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }

        public TreeNode GetNode(string nodeId)
        {
            var nodes = QueryNodes(
                $"SELECT * FROM [{_tableName}] WHERE [{_idCol}] = @nodeId",
                new OleDbParameter("@nodeId", nodeId));
            return nodes.Count > 0 ? nodes[0] : null;
        }

        public List<TreeNode> Find(string text, int maxResults)
        {
            return QueryNodes(
                $"SELECT TOP {maxResults} * FROM [{_tableName}] WHERE [{_captionCol}] LIKE @text",
                new OleDbParameter("@text", $"%{text}%"));
        }

        private List<TreeNode> QueryNodes(string sql, params OleDbParameter[] parameters)
        {
            var result = new List<TreeNode>();
            using var conn = new OleDbConnection(ConnectionString);
            conn.Open();
            using var cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddRange(parameters);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var node = new TreeNode(
                    id: reader[_idCol]?.ToString() ?? "",
                    parentId: reader[_parentIdCol]?.ToString() ?? "",
                    caption: reader[_captionCol]?.ToString() ?? ""
                );
                if (_iconKeyCol != null && reader[_iconKeyCol] != DBNull.Value)
                    node.IconKey = reader[_iconKeyCol].ToString();
                result.Add(node);
            }
            return result;
        }
    }
}
```

**Step 4: Run test — expect PASS**

**Step 5: Wire Initialize() in TreeEngine to create OleDbProvider**

Update `TreeEngine.Initialize()` to parse connection string and create an `OleDbProvider` wrapped in `CachingProviderDecorator`. The connection string format could be:
`"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=mydb.accdb;Table=tblTree;IdCol=NodeID;ParentCol=ParentID;CaptionCol=NodeText"`

**Step 6: Commit**

```bash
git commit -am "feat(engine): implement OleDbProvider for Access database trees"
```

---

### Task 1.7: COM Registration Smoke Test

**Files:**
- Create: `src/TreeEngine64/Properties/AssemblyInfo.cs` (ensure COM attributes)

**Step 1: Build in Release mode**

```bash
dotnet build src/TreeEngine64/ -c Release
```

**Step 2: Register with regasm (on Windows)**

```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase TreeEngine64.dll /tlb
```

**Step 3: Test from VBA**

Open Access, create a new module, run:
```vba
Sub TestEngine()
    Dim eng As Object
    Set eng = CreateObject("MeKo.TreeEngine")
    Debug.Print TypeName(eng)  ' Should print "TreeEngine"
End Sub
```

Expected: Prints "TreeEngine" without error.

**Step 4: Commit**

```bash
git commit -am "chore(engine): verify COM registration works"
```

---

## Phase 2 — TreeViewHost64 Visual ActiveX Control

### Task 2.1: Define Host Interfaces

**Files:**
- Create: `src/TreeViewHost64/ITreeViewHost.cs`
- Create: `src/TreeViewHost64/ITreeViewHostEvents.cs`

**Step 1: Implement the COM-visible interface**

`src/TreeViewHost64/ITreeViewHost.cs`:
```csharp
using System;
using System.Runtime.InteropServices;
using MeKo.TreeEngine;

namespace MeKo.TreeViewHost
{
    [ComVisible(true)]
    [Guid("B1B2C3D4-1111-1111-1111-000000000001")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITreeViewHost
    {
        object Engine { get; set; }
        string SelectedNodeId { get; }
        void Initialize(object engine);
        void Reload();
        void ExpandNode(string nodeId);
        void CollapseNode(string nodeId);
        void SelectNode(string nodeId);
        bool FindAndSelect(string text);
    }
}
```

`src/TreeViewHost64/ITreeViewHostEvents.cs`:
```csharp
using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeViewHost
{
    [ComVisible(true)]
    [Guid("B1B2C3D4-1111-1111-1111-000000000002")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITreeViewHostEvents
    {
        [DispId(1)] void NodeClick(string nodeId);
        [DispId(2)] void NodeDoubleClick(string nodeId);
        [DispId(3)] void BeforeExpand(string nodeId, ref bool cancel);
        [DispId(4)] void AfterExpand(string nodeId);
        [DispId(5)] void AfterCollapse(string nodeId);
        [DispId(6)] void AfterSelect(string nodeId);
        [DispId(7)] void OnError(string message);
    }
}
```

**Step 2: Commit**

```bash
git commit -am "feat(host): define ITreeViewHost and ITreeViewHostEvents COM interfaces"
```

---

### Task 2.2: Create WinForms UserControl Shell

**Files:**
- Create: `src/TreeViewHost64/TreeViewHostControl.cs`

**Step 1: Implement the basic UserControl**

```csharp
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MeKo.TreeEngine;

namespace MeKo.TreeViewHost
{
    [ComVisible(true)]
    [Guid("B1B2C3D4-2222-2222-2222-000000000001")]
    [ProgId("MeKo.TreeViewHost")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITreeViewHost))]
    [ComSourceInterfaces(typeof(ITreeViewHostEvents))]
    public class TreeViewHostControl : UserControl, ITreeViewHost
    {
        private TreeView _treeView;
        private ITreeEngine _engine;
        private const string DummyNodeKey = "__loading__";

        // COM events — delegates matching ITreeViewHostEvents signatures
        public delegate void NodeClickHandler(string nodeId);
        public delegate void NodeDoubleClickHandler(string nodeId);
        public delegate void BeforeExpandHandler(string nodeId, ref bool cancel);
        public delegate void AfterExpandHandler(string nodeId);
        public delegate void AfterCollapseHandler(string nodeId);
        public delegate void AfterSelectHandler(string nodeId);
        public delegate void OnErrorHandler(string message);

        public event NodeClickHandler NodeClick;
        public event NodeDoubleClickHandler NodeDoubleClick;
        public event BeforeExpandHandler BeforeExpand;
        public event AfterExpandHandler AfterExpand;
        public event AfterCollapseHandler AfterCollapse;
        public event AfterSelectHandler AfterSelect;
        public event OnErrorHandler OnError;

        public TreeViewHostControl()
        {
            _treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false
            };
            Controls.Add(_treeView);

            _treeView.BeforeExpand += TreeView_BeforeExpand;
            _treeView.AfterCollapse += TreeView_AfterCollapse;
            _treeView.AfterSelect += TreeView_AfterSelect;
            _treeView.NodeMouseClick += TreeView_NodeMouseClick;
            _treeView.NodeMouseDoubleClick += TreeView_NodeMouseDoubleClick;
        }

        // ITreeViewHost implementation
        public object Engine
        {
            get => _engine;
            set => Initialize(value);
        }

        public string SelectedNodeId =>
            _treeView.SelectedNode?.Name ?? "";

        public void Initialize(object engine)
        {
            _engine = (ITreeEngine)engine;
            Reload();
        }

        public void Reload()
        {
            // Implemented in Task 2.3
        }

        public void ExpandNode(string nodeId)
        {
            var node = FindTreeNode(nodeId);
            node?.Expand();
        }

        public void CollapseNode(string nodeId)
        {
            var node = FindTreeNode(nodeId);
            node?.Collapse();
        }

        public void SelectNode(string nodeId)
        {
            var node = FindTreeNode(nodeId);
            if (node != null)
                _treeView.SelectedNode = node;
        }

        public bool FindAndSelect(string text)
        {
            // Implemented in Task 2.5
            return false;
        }

        private TreeNode FindTreeNode(string nodeId)
        {
            var found = _treeView.Nodes.Find(nodeId, searchAllChildren: true);
            return found.Length > 0 ? found[0] : null;
        }

        // Event handlers — stubs, wired in Task 2.3/2.4
        private void TreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e) { }
        private void TreeView_AfterCollapse(object sender, TreeViewEventArgs e) { }
        private void TreeView_AfterSelect(object sender, TreeViewEventArgs e) { }
        private void TreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e) { }
        private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e) { }
    }
}
```

**Step 2: Build to verify compilation**

```bash
dotnet build src/TreeViewHost64/
```

**Step 3: Commit**

```bash
git commit -am "feat(host): create TreeViewHostControl UserControl shell"
```

---

### Task 2.3: Implement Tree Loading (Root + Lazy Children)

**Files:**
- Modify: `src/TreeViewHost64/TreeViewHostControl.cs`

**Step 1: Implement Reload() — load root nodes**

```csharp
public void Reload()
{
    _treeView.BeginUpdate();
    try
    {
        _treeView.Nodes.Clear();
        if (_engine == null) return;

        var roots = _engine.GetRootNodes();
        for (int i = 1; i <= roots.Count; i++)
        {
            var data = roots[i];
            var treeNode = CreateVisualNode(data);
            _treeView.Nodes.Add(treeNode);
        }
    }
    catch (Exception ex)
    {
        OnError?.Invoke(ex.Message);
    }
    finally
    {
        _treeView.EndUpdate();
    }
}

private TreeNode CreateVisualNode(ITreeNode data)
{
    var treeNode = new TreeNode(data.Caption)
    {
        Name = data.Id,
        Tag = data.Id,
        ImageKey = data.IconKey,
        SelectedImageKey = data.IconKey
    };

    // Add dummy child if engine says this node has children
    if (_engine.HasChildren(data.Id))
    {
        treeNode.Nodes.Add(DummyNodeKey, "Loading...");
    }

    return treeNode;
}
```

**Step 2: Implement BeforeExpand — lazy load children**

```csharp
private void TreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
{
    try
    {
        var nodeId = e.Node.Name;

        // Raise cancelable event
        bool cancel = false;
        BeforeExpand?.Invoke(nodeId, ref cancel);
        if (cancel) { e.Cancel = true; return; }

        // Check if children need loading (dummy node present)
        if (e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Name == DummyNodeKey)
        {
            e.Node.Nodes.Clear();
            _treeView.BeginUpdate();
            try
            {
                var children = _engine.GetChildren(nodeId);
                for (int i = 1; i <= children.Count; i++)
                {
                    e.Node.Nodes.Add(CreateVisualNode(children[i]));
                }
            }
            finally
            {
                _treeView.EndUpdate();
            }
        }

        AfterExpand?.Invoke(nodeId);
    }
    catch (Exception ex)
    {
        OnError?.Invoke(ex.Message);
    }
}
```

**Step 3: Build and verify**

```bash
dotnet build src/TreeViewHost64/
```

**Step 4: Commit**

```bash
git commit -am "feat(host): implement Reload and lazy-loading BeforeExpand"
```

---

### Task 2.4: Wire Remaining Events

**Files:**
- Modify: `src/TreeViewHost64/TreeViewHostControl.cs`

**Step 1: Implement event handlers**

```csharp
private void TreeView_AfterCollapse(object sender, TreeViewEventArgs e)
{
    AfterCollapse?.Invoke(e.Node.Name);
}

private void TreeView_AfterSelect(object sender, TreeViewEventArgs e)
{
    AfterSelect?.Invoke(e.Node.Name);
}

private void TreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
{
    NodeClick?.Invoke(e.Node.Name);
}

private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
{
    NodeDoubleClick?.Invoke(e.Node.Name);
}
```

**Step 2: Build and verify**

**Step 3: Commit**

```bash
git commit -am "feat(host): wire AfterCollapse, AfterSelect, NodeClick, NodeDoubleClick events"
```

---

### Task 2.5: Implement FindAndSelect

**Files:**
- Modify: `src/TreeViewHost64/TreeViewHostControl.cs`

**Step 1: Implement FindAndSelect**

```csharp
public bool FindAndSelect(string text)
{
    if (_engine == null) return false;

    try
    {
        var results = _engine.Find(text, 1);
        if (results.Count == 0) return false;

        var nodeId = results[1].Id;

        // Expand parent chain to make node visible
        ExpandParentChain(nodeId);

        // Select the node
        var treeNode = FindTreeNode(nodeId);
        if (treeNode != null)
        {
            _treeView.SelectedNode = treeNode;
            treeNode.EnsureVisible();
            return true;
        }
    }
    catch (Exception ex)
    {
        OnError?.Invoke(ex.Message);
    }
    return false;
}

private void ExpandParentChain(string nodeId)
{
    var node = _engine.GetNode(nodeId);
    if (node == null || string.IsNullOrEmpty(node.ParentId)) return;

    // Build parent chain
    var chain = new System.Collections.Generic.Stack<string>();
    var current = node;
    while (current != null && !string.IsNullOrEmpty(current.ParentId))
    {
        chain.Push(current.ParentId);
        current = _engine.GetNode(current.ParentId);
    }

    // Expand from root down
    foreach (var parentId in chain)
    {
        var parentNode = FindTreeNode(parentId);
        if (parentNode != null && !parentNode.IsExpanded)
            parentNode.Expand();
    }
}
```

**Step 2: Build and verify**

**Step 3: Commit**

```bash
git commit -am "feat(host): implement FindAndSelect with parent chain expansion"
```

---

### Task 2.6: Optional Features — Checkboxes, ImageList, Context Menu

**Files:**
- Modify: `src/TreeViewHost64/ITreeViewHost.cs` (add properties)
- Modify: `src/TreeViewHost64/TreeViewHostControl.cs`

**Step 1: Add optional feature properties to interface**

```csharp
// Add to ITreeViewHost
bool CheckBoxes { get; set; }
void SetImageList(object imageListHandle); // stdole.IPictureDisp or handle
```

**Step 2: Implement in control**

```csharp
public bool CheckBoxes
{
    get => _treeView.CheckBoxes;
    set => _treeView.CheckBoxes = value;
}
```

**Step 3: Build and verify**

**Step 4: Commit**

```bash
git commit -am "feat(host): add CheckBoxes property and ImageList support"
```

---

### Task 2.7: COM Registration & Access Form Test

**Step 1: Build Release**

```bash
dotnet build src/TreeViewHost64/ -c Release
```

**Step 2: Register both DLLs (on Windows)**

```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase TreeEngine64.dll /tlb
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase TreeViewHost64.dll /tlb
```

**Step 3: Test in Access**

1. Open Access database
2. Insert ActiveX control on a form → look for "MeKo.TreeViewHost"
3. In form's VBA module:

```vba
Dim WithEvents tvHost As TreeViewHost64.TreeViewHostControl
Dim eng As Object

Private Sub Form_Load()
    Set eng = CreateObject("MeKo.TreeEngine")
    ' For now, engine with no data — just verifying COM works
    Set tvHost = Me.TreeViewHost1.Object
    tvHost.Initialize eng
End Sub

Private Sub tvHost_AfterSelect(ByVal nodeId As String)
    Debug.Print "Selected: " & nodeId
End Sub
```

**Step 4: Commit**

```bash
git commit -am "test: verify COM registration and Access form hosting"
```

---

## Phase 3 — Access Integration

### Task 3.1: Create VBA Wrapper Module

**Files:**
- Create: `vba/modTreeCompat.bas`

**Step 1: Write the VBA compatibility module**

```vba
Option Compare Database
Option Explicit

' modTreeCompat — Drop-in helpers for MeKo TreeView
' Usage:
'   Tree_Init Me.ctlTree, "Provider=Microsoft.ACE.OLEDB.16.0;..."
'   Tree_SelectByKey Me.ctlTree, "42"
'   Tree_Refresh Me.ctlTree

Public Sub Tree_Init(ctl As Object, connectionString As String, _
                     Optional tableName As String = "tblTreeNodes", _
                     Optional idCol As String = "NodeID", _
                     Optional parentCol As String = "ParentID", _
                     Optional captionCol As String = "NodeText")
    Dim eng As Object
    Set eng = CreateObject("MeKo.TreeEngine")
    eng.Initialize connectionString
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
```

**Step 2: Commit**

```bash
git commit -am "feat(vba): add modTreeCompat.bas compatibility wrapper"
```

---

### Task 3.2: Create Demo Access Form

**Files:**
- Create: `demo/DemoTreeView.accdb` (manual — Access form)
- Create: `vba/Form_frmTreeDemo.cls`

**Step 1: Write the demo form's VBA class**

```vba
Option Compare Database
Option Explicit

Dim WithEvents tvHost As TreeViewHost64.TreeViewHostControl

Private Sub Form_Load()
    Set tvHost = Me.ctlTreeView.Object

    Dim eng As Object
    Set eng = CreateObject("MeKo.TreeEngine")

    ' Point to demo table in this database
    eng.Initialize "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & CurrentDb.Name

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
```

**Step 2: Create demo table structure**

The demo Access DB should contain a table `tblTreeNodes`:

| Column | Type | Description |
|---|---|---|
| NodeID | Text(50) PK | Unique node ID |
| ParentID | Text(50) | Parent node ID (empty for root) |
| NodeText | Text(255) | Display text |
| IconKey | Text(50) | Optional icon key |

Insert sample data:
```sql
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1', '', 'Company');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.1', '1', 'Engineering');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.2', '1', 'Sales');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.1.1', '1.1', 'Backend Team');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.1.2', '1.1', 'Frontend Team');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.2.1', '1.2', 'DACH Region');
INSERT INTO tblTreeNodes (NodeID, ParentID, NodeText) VALUES ('1.2.2', '1.2', 'International');
```

**Step 3: Commit**

```bash
git commit -am "feat(demo): add demo form VBA and sample table schema"
```

---

## Phase 4 — Installer & Deployment

### Task 4.1: Create WiX Installer Project

**Files:**
- Create: `installer/MeKoTreeView.wixproj`
- Create: `installer/Product.wxs`

**Step 1: Install WiX Toolset**

```bash
dotnet tool install --global wix
```

**Step 2: Write the WiX product definition**

`installer/Product.wxs`:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">
  <Package Name="MeKo TreeView for Access"
           Manufacturer="MeKo-Tech"
           Version="1.0.0.0"
           UpgradeCode="PUT-REAL-GUID-HERE">

    <MajorUpgrade DowngradeErrorMessage="A newer version is already installed." />
    <MediaTemplate EmbedCab="yes" />

    <Feature Id="ProductFeature" Title="MeKo TreeView" Level="1">
      <ComponentGroupRef Id="TreeEngineComponents" />
      <ComponentGroupRef Id="TreeViewHostComponents" />
    </Feature>

    <!-- Custom actions for regasm registration -->
    <!-- TODO: Add CustomAction for regasm /codebase during install -->
    <!-- TODO: Add CustomAction for regasm /unregister during uninstall -->
  </Package>
</Wix>
```

**Step 3: Build installer**

```bash
dotnet build installer/MeKoTreeView.wixproj -c Release
```

**Step 4: Test install/uninstall on a clean Windows VM**

**Step 5: Commit**

```bash
git commit -am "chore(installer): scaffold WiX installer for COM registration"
```

---

### Task 4.2: Add Registration Custom Actions

**Files:**
- Modify: `installer/Product.wxs`

**Step 1: Add regasm custom actions**

The installer must run `regasm /codebase` for both DLLs during install and `regasm /unregister` during uninstall. This ensures proper COM registration in the 64-bit registry hive.

**Step 2: Test on clean machine**

- Install MSI
- Open Access → CreateObject("MeKo.TreeEngine") works
- Uninstall MSI
- CreateObject fails (as expected)

**Step 3: Commit**

```bash
git commit -am "feat(installer): add regasm registration custom actions"
```

---

## Phase 5 — Testing

### Task 5.1: Unit Test Suite Completion

**Files:**
- Modify: `tests/TreeEngine64.Tests/*.cs`

**Step 1: Add edge case tests**

- [ ] Empty tree (no root nodes)
- [ ] Node with no children
- [ ] Find with no results
- [ ] Find with special characters in search text
- [ ] GetNode with null/empty ID
- [ ] Very long caption text
- [ ] Concurrent access to caching provider

**Step 2: Run all tests**

```bash
dotnet test tests/TreeEngine64.Tests/ -v n
```

Expected: All pass.

**Step 3: Commit**

```bash
git commit -am "test: add edge case tests for TreeEngine"
```

---

### Task 5.2: Integration Test with Access Database

**Files:**
- Create: `tests/TreeEngine64.Tests/Integration/OleDbIntegrationTests.cs`

**Step 1: Write integration tests** (marked `[Category("Integration")]`)

- [ ] OleDbProvider loads root nodes from real .accdb file
- [ ] OleDbProvider loads children
- [ ] OleDbProvider handles empty table
- [ ] OleDbProvider handles missing table (error case)
- [ ] Full round-trip: engine → provider → cache → retrieve

**Step 2: Run integration tests (Windows only)**

```bash
dotnet test tests/TreeEngine64.Tests/ --filter Category=Integration -v n
```

**Step 3: Commit**

```bash
git commit -am "test: add OleDb integration tests"
```

---

### Task 5.3: Access-Specific Manual Test Checklist

Create `docs/test-checklist.md`:

- [ ] Insert control on Access form via form designer
- [ ] Control renders and shows tree on form open
- [ ] Close and reopen form — control still works
- [ ] Close and reopen database — control still works
- [ ] Multiple instances of control on different forms
- [ ] Compile database to ACCDE — control still works
- [ ] Expand node — children load correctly
- [ ] Collapse and re-expand — children still present
- [ ] Click node — event fires in VBA
- [ ] Double-click node — event fires in VBA
- [ ] Search and select — node found and highlighted
- [ ] Tree with 1000+ nodes — no visible lag on expand
- [ ] Tree with 10000+ nodes — lazy loading keeps UI responsive
- [ ] Rapid expand/collapse — no crashes or visual glitches

**Commit:**

```bash
git commit -am "docs: add Access-specific manual test checklist"
```

---

## Summary of All Tasks

| Phase | Task | Description |
|---|---|---|
| 0 | 0.1 | Language evaluation & decision (C#) |
| 0 | 0.2 | Solution structure & project scaffolding |
| 0 | 0.3 | COM contract definition |
| 1 | 1.1 | COM interfaces (ITreeNode, ITreeNodeCollection, ITreeEngine) |
| 1 | 1.2 | TreeNode implementation |
| 1 | 1.3 | TreeNodeCollection implementation |
| 1 | 1.4 | TreeEngine core with InMemoryProvider |
| 1 | 1.5 | Caching layer |
| 1 | 1.6 | OleDb data provider |
| 1 | 1.7 | COM registration smoke test |
| 2 | 2.1 | Host COM interfaces |
| 2 | 2.2 | WinForms UserControl shell |
| 2 | 2.3 | Tree loading + lazy expand |
| 2 | 2.4 | Wire remaining events |
| 2 | 2.5 | FindAndSelect with parent chain expansion |
| 2 | 2.6 | Optional features (checkboxes, images) |
| 2 | 2.7 | COM registration & Access form test |
| 3 | 3.1 | VBA compatibility wrapper module |
| 3 | 3.2 | Demo Access form |
| 4 | 4.1 | WiX installer scaffolding |
| 4 | 4.2 | Registration custom actions |
| 5 | 5.1 | Unit test edge cases |
| 5 | 5.2 | Integration tests |
| 5 | 5.3 | Manual test checklist |
