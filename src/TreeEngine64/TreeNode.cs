using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine;

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
