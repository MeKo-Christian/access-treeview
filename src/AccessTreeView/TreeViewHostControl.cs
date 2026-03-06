using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.ComponentModel;
using Access.TreeEngine;
using WinTreeNode = System.Windows.Forms.TreeNode;

namespace MeKo.TreeViewHost;

[ComVisible(true)]
[Guid("B1B2C3D4-2222-2222-2222-000000000001")]
[ProgId("MeKo.TreeViewHost")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(ITreeViewHost))]
[ComSourceInterfaces(typeof(ITreeViewHostEvents))]
public class TreeViewHostControl : UserControl, ITreeViewHost
{
    private readonly TreeView _treeView;
    private ITreeEngine _engine;
    private const string DummyNodeKey = "__loading__";

    // COM event delegates matching ITreeViewHostEvents signatures
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

    // --- ITreeViewHost Properties ---

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    [Browsable(false)]
    public object Engine
    {
        get => _engine;
        set => Initialize(value);
    }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    [Browsable(false)]
    public string SelectedNodeId => _treeView.SelectedNode?.Name ?? "";

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
    public bool CheckBoxes
    {
        get => _treeView.CheckBoxes;
        set => _treeView.CheckBoxes = value;
    }

    // --- ITreeViewHost Methods ---

    public void Initialize(object engine)
    {
        _engine = (ITreeEngine)engine;
        Reload();
    }

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

    // --- Event Handlers ---

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

    // --- Helpers ---

    private WinTreeNode CreateVisualNode(ITreeNode data)
    {
        var treeNode = new WinTreeNode(data.Caption)
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

    private WinTreeNode FindTreeNode(string nodeId)
    {
        var found = _treeView.Nodes.Find(nodeId, searchAllChildren: true);
        return found.Length > 0 ? found[0] : null;
    }

    private void ExpandParentChain(string nodeId)
    {
        var node = _engine.GetNode(nodeId);
        if (node == null || string.IsNullOrEmpty(node.ParentId)) return;

        // Build parent chain
        var chain = new Stack<string>();
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
}
