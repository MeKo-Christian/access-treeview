using NUnit.Framework;
using Access.TreeEngine;

namespace AccessTreeEngine.Tests;

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

    [Test]
    public void TreeNode_Null_ParentId_Becomes_Empty()
    {
        var node = new TreeNode("1", null, "Root");
        Assert.That(node.ParentId, Is.EqualTo(""));
    }

    [Test]
    public void TreeNode_Null_Caption_Becomes_Empty()
    {
        var node = new TreeNode("1", "", null);
        Assert.That(node.Caption, Is.EqualTo(""));
    }
}
