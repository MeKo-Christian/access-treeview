using NUnit.Framework;
using MeKo.TreeEngine;
using System.Collections.Generic;

namespace TreeEngine64.Tests;

[TestFixture]
public class FullRoundTripTests
{
    [Test]
    public void Full_RoundTrip_With_Caching()
    {
        // Setup: engine with caching decorator wrapping in-memory provider
        var nodes = new List<TreeNode>
        {
            new TreeNode("1", "", "Company"),
            new TreeNode("1.1", "1", "Engineering"),
            new TreeNode("1.2", "1", "Sales"),
            new TreeNode("1.1.1", "1.1", "Backend"),
            new TreeNode("1.1.2", "1.1", "Frontend"),
            new TreeNode("1.2.1", "1.2", "DACH"),
            new TreeNode("1.2.2", "1.2", "International"),
        };
        var inner = new InMemoryProvider(nodes);
        var caching = new CachingProviderDecorator(inner);
        var engine = new TreeEngine();
        engine.SetProvider(caching);

        // Get roots
        var roots = engine.GetRootNodes();
        Assert.That(roots.Count, Is.EqualTo(1));
        Assert.That(roots[1].Caption, Is.EqualTo("Company"));

        // Expand first level
        var level1 = engine.GetChildren("1");
        Assert.That(level1.Count, Is.EqualTo(2));

        // Expand second level
        var level2 = engine.GetChildren("1.1");
        Assert.That(level2.Count, Is.EqualTo(2));
        Assert.That(level2[1].Caption, Is.EqualTo("Backend"));

        // Verify caching: second call returns same data
        var level2Again = engine.GetChildren("1.1");
        Assert.That(level2Again.Count, Is.EqualTo(level2.Count));

        // Invalidate and re-fetch
        engine.Invalidate("1.1");
        var level2Refreshed = engine.GetChildren("1.1");
        Assert.That(level2Refreshed.Count, Is.EqualTo(2));

        // Search
        var found = engine.Find("Backend");
        Assert.That(found.Count, Is.EqualTo(1));
        Assert.That(found[1].Id, Is.EqualTo("1.1.1"));

        // HasChildren
        Assert.That(engine.HasChildren("1"), Is.True);
        Assert.That(engine.HasChildren("1.2.1"), Is.False);

        // GetNode
        var node = engine.GetNode("1.2.2");
        Assert.That(node.Caption, Is.EqualTo("International"));
    }

    [Test]
    public void Reload_Clears_All_Caches()
    {
        var nodes = new List<TreeNode>
        {
            new TreeNode("1", "", "Root"),
            new TreeNode("1.1", "1", "Child"),
        };
        var inner = new InMemoryProvider(nodes);
        var caching = new CachingProviderDecorator(inner);
        var engine = new TreeEngine();
        engine.SetProvider(caching);

        // Populate cache
        engine.GetChildren("1");

        // Reload should clear cache
        engine.Reload();

        // Should work fine (re-fetches from provider)
        var children = engine.GetChildren("1");
        Assert.That(children.Count, Is.EqualTo(1));
    }
}
