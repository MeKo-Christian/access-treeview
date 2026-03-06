using NUnit.Framework;
using Access.TreeEngine;
using System.Collections.Generic;

namespace AccessTreeEngine.Tests;

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

    [Test]
    public void Throws_When_No_Provider_Set()
    {
        var uninitEngine = new TreeEngine();
        Assert.Throws<System.InvalidOperationException>(() => uninitEngine.GetRootNodes());
    }
}
