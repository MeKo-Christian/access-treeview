using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeViewHost;

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
