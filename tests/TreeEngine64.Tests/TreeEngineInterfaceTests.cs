using NUnit.Framework;
using MeKo.TreeEngine;

namespace TreeEngine64.Tests;

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
        Assert.That(type.GetMethod("Invalidate"), Is.Not.Null);
        Assert.That(type.GetMethod("Reload"), Is.Not.Null);
    }

    [Test]
    public void ITreeNodeCollection_Has_Count_Property()
    {
        var type = typeof(ITreeNodeCollection);
        Assert.That(type.GetProperty("Count"), Is.Not.Null);
    }
}
