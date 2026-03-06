using System;
using System.Runtime.InteropServices;

namespace Access.TreeEngine;

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
