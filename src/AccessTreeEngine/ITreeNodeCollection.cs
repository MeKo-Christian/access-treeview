using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace Access.TreeEngine;

[ComVisible(true)]
[Guid("A1B2C3D4-1111-1111-1111-000000000002")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface ITreeNodeCollection
{
    int Count { get; }
    ITreeNode this[int index] { get; }

    [DispId(-4)]
    IEnumerator GetEnumerator();
}
