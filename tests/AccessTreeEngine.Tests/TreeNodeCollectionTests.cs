using NUnit.Framework;
using Access.TreeEngine;
using System.Collections.Generic;

namespace AccessTreeEngine.Tests;

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

    [Test]
    public void Null_List_Creates_Empty_Collection()
    {
        var coll = new TreeNodeCollection(null);
        Assert.That(coll.Count, Is.EqualTo(0));
    }
}
