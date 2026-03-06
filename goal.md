## Target outcome

In Access (64-bit) you can:

* Place **Access TreeView** on a form (like an OCX control)
* Bind it to your tree data source
* Handle events in VBA (`NodeClick`, `Expand`, `Collapse`, etc.)
* Keep your tree logic/data access out of the form and inside a **64-bit COM DLL**

---

# High-level architecture

### Component 1 — AccessTreeEngine (COM in-proc DLL)

**Purpose:** data + behavior

* Builds nodes, lazy-loads children, applies filters/search
* Exposes a VBA-friendly API (`GetChildren`, `Expand`, `Collapse`, `Find`, …)
* Optional persistence to tables/queries

**Tech choice:**

* **C++/ATL** (best binary COM discipline) **or** **C# (.NET) COM-visible class library** (faster to build).
  Given you’re also building a .NET UI control, **C# is usually the fastest overall** (one ecosystem).

### Component 2 — TreeViewHost64 (Visual ActiveX control)

**Purpose:** UI only

* A **WinForms UserControl** wrapping `System.Windows.Forms.TreeView`
* Exposed to COM so Access can host it on a form (ActiveX-like)
* Talks to `AccessTreeEngine` via COM
* Raises VBA events

---

# Implementation plan (end-to-end)

## Phase 0 — Compatibility “contract” (1–2 days)

Deliver a short spec so the new control is “as close as possible” to the old one.

1. **List existing VBA usage patterns**

* How nodes are added today (`Nodes.Add`, `Key`, `Text`, `Tag`, images)
* Which events are handled (click, dblclick, expand/collapse, label edit)
* Any drag & drop / checkboxes / multiselect

2. Define your **public API surface** that VBA will call in the new control:

* Must-have methods/properties
* Must-have events

3. Define data feeding strategy:

* “push nodes from VBA” (old style) vs “engine loads from DB”
* For large trees: **lazy-loading** required? (recommended)

**Deliverable:** “TreeView COM Contract” (method list + event list + node fields).

---

## Phase 1 — AccessTreeEngine COM DLL (2–6 days)

### 1.1 Object model

Keep it simple and VBA-friendly.

**Interfaces**

* `ITreeEngine`

  * `Initialize(connectionStringOrDsn As String, Optional context As Variant)`
  * `GetRootNodes() As ITreeNodeCollection`
  * `GetChildren(nodeId As String) As ITreeNodeCollection`
  * `HasChildren(nodeId As String) As Boolean`
  * `Find(text As String, Optional maxResults As Long = 100) As ITreeNodeCollection`
  * `GetNode(nodeId As String) As ITreeNode`
  * `Invalidate(nodeId As String)` / `Reload()`
* `ITreeNode`

  * `Id As String`
  * `ParentId As String`
  * `Caption As String`
  * `IconKey As String` (or `IconIndex As Long`)
  * `Tag As Variant`
* `ITreeNodeCollection`

  * `_NewEnum` for `For Each`
  * `Count`, `Item(index)`

### 1.2 Data access strategy

Pick one:

* **Engine queries DB directly** (recommended for performance + clean Access forms)
* or engine calls back to VBA “provider” (more flexible but more complex)

**Recommended:** engine loads from DB using:

* ADO/OleDb/ODBC (if C#) or ODBC/ADO (if C++)

Lazy-loading design:

* only load children when UI expands a node
* cache children by nodeId (with invalidation)

### 1.3 Versioning rules

* Stable ProgID (e.g. `Access.TreeEngine`)
* New functionality: add new interfaces rather than breaking old signatures

**Deliverables:**

* `AccessTreeEngine.dll` (registered COM server)
* Small VBA test module proving `CreateObject("Access.TreeEngine")` works

---

## Phase 2 — TreeViewHost64 Visual ActiveX (4–10 days)

This is the key Track 2 work.

### 2.1 UI control features (baseline)

Implement a WinForms UserControl with:

* embedded `TreeView`
* support:

  * icons (ImageList)
  * context menu hooks
  * optional checkboxes
  * optional label edit
* node object mapping:

  * `TreeNode.Name` = nodeId
  * `TreeNode.Text` = caption
  * `TreeNode.Tag` = (optional) cached `ITreeNode` or nodeId

### 2.2 COM exposure so Access can host it

Expose the WinForms control via COM:

* `[ComVisible(true)]`, `[Guid]`, `[ProgId]`
* Implement a COM-visible interface `ITreeViewHost`
* Ensure it can be inserted on an Access form (registered correctly)

**Public API of the control (VBA-facing)**

* `Engine` property (object)
* `Initialize(engine As Object)`
* `Reload()`
* `ExpandNode(nodeId As String)`
* `SelectNode(nodeId As String)`
* `FindAndSelect(text As String) As Boolean`
* `SelectedNodeId As String` (get)

**Events raised to VBA**

* `NodeClick(nodeId As String)`
* `NodeDoubleClick(nodeId As String)`
* `BeforeExpand(nodeId As String, Cancel As Boolean)`
* `AfterExpand(nodeId As String)`
* `AfterCollapse(nodeId As String)`
* `AfterSelect(nodeId As String)`
* `Error(message As String)`

Eventing in COM typically uses a **source interface** (`[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]`) so VBA can `WithEvents`.

### 2.3 Wiring control ↔ engine

The UI control never “builds” business logic. It:

* calls `engine.GetRootNodes()` at load
* on expand:

  * calls `engine.GetChildren(nodeId)`
  * creates visual nodes
* on select:

  * updates `SelectedNodeId`
  * raises event to VBA

### 2.4 Performance / UX details

* Use a dummy child (“Loading…”) to show expand glyph before children are loaded
* Load children on-demand
* Keep calls to COM engine minimal
* Add `BeginUpdate/EndUpdate` around bulk node insert

**Deliverables:**

* `TreeViewHost64.dll` registered and insertable in Access form designer
* Demo Access form showing the control running off dummy data

---

## Phase 3 — Access integration (1–3 days)

### 3.1 Reference & insertion

* Add the control to a form (ActiveX)
* Provide a minimal VBA wrapper module so your app code stays clean

### 3.2 VBA glue template

* Instantiate engine in `Form_Load`
* Assign to control
* Handle `WithEvents` for user interactions
* Provide compatibility helpers to mimic your previous tree API (if needed)

**Deliverable:** drop-in `modTreeCompat.bas` that offers functions like:

* `Tree_Init(form.TreeControl, "connection string")`
* `Tree_SelectByKey(...)`
* `Tree_Refresh(...)`

---

## Phase 4 — Installer & deployment (2–5 days)

You’ll need a proper installer. For 64-bit Access you must register 64-bit COM components.

### 4.1 Packaging

Package **both**:

* AccessTreeEngine (COM in-proc)
* TreeViewHost64 (COM-visible control)

### 4.2 Registration

* **Preferred:** MSI (WiX Toolset)

  * registers COM class + type library
  * sets correct registry keys in 64-bit hive
  * installs to Program Files
* Dev registration can use `regasm` (for .NET), but production should not rely on manual steps.

### 4.3 Optional dual-architecture support (if needed later)

If you still have users on 32-bit Access:

* build x86 versions too
* ship two MSIs or one bootstrapper that installs the correct bitness

**Deliverable:** signed installer, silent install option, and an uninstall clean-up checklist.

---

## Phase 5 — Testing matrix (ongoing; 2–4 days initial)

### 5.1 Functional tests

* Expand/collapse
* Selection + events
* Search
* Label edit (if needed)
* Context menu actions

### 5.2 Access-specific tests

* Form designer load/unload
* Reopen database
* Multiple instances of the control on different forms
* Stability when Access is compiled to ACCDE

### 5.3 Performance tests

* Large trees (10k+ nodes with lazy loading)
* Rapid expand/collapse
* Search responsiveness

**Deliverable:** test checklist + sample dataset generator.

---

# Practical recommendations (to reduce risk)

1. **Don’t try to replicate MSCOMCTL API 1:1 internally**. Instead:

   * keep your VBA touchpoints similar,
   * provide a small VBA compatibility layer for any “Nodes.Add”-style code you must keep.
2. **Lazy-load** from day one. Full materialization kills performance in Access UIs.
3. Use an installer early (even in dev) so “it works on my machine” doesn’t become a time sink.
