using NUnit.Framework;
using Access.TreeEngine;
using System.Collections.Generic;

namespace AccessTreeEngine.Tests;

[TestFixture]
public class EdgeCaseTests
{
    [Test]
    public void Empty_Tree_Returns_No_Root_Nodes()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>()));
        var roots = engine.GetRootNodes();
        Assert.That(roots.Count, Is.EqualTo(0));
    }

    [Test]
    public void Find_With_No_Results_Returns_Empty_Collection()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Root")
        }));
        var results = engine.Find("nonexistent");
        Assert.That(results.Count, Is.EqualTo(0));
    }

    [Test]
    public void Find_With_Special_Characters()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Node (special) [chars]")
        }));
        var results = engine.Find("(special)");
        Assert.That(results.Count, Is.EqualTo(1));
    }

    [Test]
    public void GetNode_With_Empty_String_Returns_Null()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Root")
        }));
        var node = engine.GetNode("");
        Assert.That(node, Is.Null);
    }

    [Test]
    public void GetNode_With_Null_Returns_Null()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Root")
        }));
        var node = engine.GetNode(null);
        Assert.That(node, Is.Null);
    }

    [Test]
    public void Very_Long_Caption()
    {
        var longCaption = new string('A', 10000);
        var node = new TreeNode("1", "", longCaption);
        Assert.That(node.Caption, Has.Length.EqualTo(10000));
    }

    [Test]
    public void TreeNodeCollection_OneBasedIndex_Throws_On_Zero()
    {
        var coll = new TreeNodeCollection(new List<TreeNode>
        {
            new TreeNode("1", "", "First")
        });
        Assert.Throws<System.ArgumentOutOfRangeException>(() =>
        {
            var _ = coll[0];
        });
    }

    [Test]
    public void TreeNodeCollection_Throws_On_Index_Beyond_Count()
    {
        var coll = new TreeNodeCollection(new List<TreeNode>
        {
            new TreeNode("1", "", "First")
        });
        Assert.Throws<System.ArgumentOutOfRangeException>(() =>
        {
            var _ = coll[2];
        });
    }

    [Test]
    public void HasChildren_For_Nonexistent_Node_Returns_False()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Root")
        }));
        Assert.That(engine.HasChildren("999"), Is.False);
    }

    [Test]
    public void Find_Case_Insensitive()
    {
        var engine = new TreeEngine();
        engine.SetProvider(new InMemoryProvider(new List<TreeNode>
        {
            new TreeNode("1", "", "Hello World")
        }));
        var results = engine.Find("hello world");
        Assert.That(results.Count, Is.EqualTo(1));
    }

    [Test]
    public void Caching_Invalidate_Nonexistent_Key_Does_Not_Throw()
    {
        var inner = new InMemoryProvider(new List<TreeNode>());
        var caching = new CachingProviderDecorator(inner);
        Assert.DoesNotThrow(() => caching.Invalidate("nonexistent"));
    }

    [Test]
    public void Engine_Reload_With_Caching_Provider()
    {
        var engine = new TreeEngine();
        var nodes = new List<TreeNode> { new TreeNode("1", "", "Root") };
        var inner = new InMemoryProvider(nodes);
        var caching = new CachingProviderDecorator(inner);
        engine.SetProvider(caching);

        // Load and cache
        engine.GetRootNodes();
        engine.GetChildren("1");

        // Reload should not throw
        Assert.DoesNotThrow(() => engine.Reload());
    }
}
