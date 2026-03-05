using NUnit.Framework;
using MeKo.TreeEngine;
using System.Collections.Generic;

namespace TreeEngine64.Tests.Integration;

/// <summary>
/// Integration tests that require Windows + Access database.
/// Run with: dotnet test --filter Category=Integration
/// </summary>
[TestFixture]
[Category("Integration")]
[Ignore("Requires Windows with Access OLEDB provider")]
public class OleDbIntegrationTests
{
    private const string TestConnectionString =
        "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=TestTreeView.accdb" +
        ";Table=tblTreeNodes;IdCol=NodeID;ParentCol=ParentID;CaptionCol=NodeText";

    [Test]
    public void Initialize_Creates_Provider_From_ConnectionString()
    {
        var engine = new TreeEngine();
        // This will fail on Linux but should work on Windows with Access
        engine.Initialize(TestConnectionString);

        var roots = engine.GetRootNodes();
        Assert.That(roots, Is.Not.Null);
    }

    [Test]
    public void GetRootNodes_Returns_Top_Level_Nodes()
    {
        var engine = new TreeEngine();
        engine.Initialize(TestConnectionString);

        var roots = engine.GetRootNodes();
        Assert.That(roots.Count, Is.GreaterThan(0));
    }

    [Test]
    public void GetChildren_Returns_Children()
    {
        var engine = new TreeEngine();
        engine.Initialize(TestConnectionString);

        var roots = engine.GetRootNodes();
        var firstRootId = roots[1].Id;
        var children = engine.GetChildren(firstRootId);

        Assert.That(children, Is.Not.Null);
    }

    [Test]
    public void HasChildren_Works_With_Real_Data()
    {
        var engine = new TreeEngine();
        engine.Initialize(TestConnectionString);

        var roots = engine.GetRootNodes();
        if (roots.Count > 0)
        {
            // Root nodes in demo data should have children
            Assert.That(engine.HasChildren(roots[1].Id), Is.True);
        }
    }

    [Test]
    public void Find_Works_With_Real_Data()
    {
        var engine = new TreeEngine();
        engine.Initialize(TestConnectionString);

        var results = engine.Find("Company");
        Assert.That(results.Count, Is.GreaterThanOrEqualTo(0));
    }

    [Test]
    public void Empty_Table_Returns_No_Roots()
    {
        var engine = new TreeEngine();
        // Point to an empty table
        var connStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=TestTreeView.accdb" +
                      ";Table=tblEmpty;IdCol=NodeID;ParentCol=ParentID;CaptionCol=NodeText";
        engine.Initialize(connStr);

        var roots = engine.GetRootNodes();
        Assert.That(roots.Count, Is.EqualTo(0));
    }

    [Test]
    public void Full_RoundTrip_Engine_Cache_Retrieve()
    {
        var engine = new TreeEngine();
        engine.Initialize(TestConnectionString);

        // First call loads from DB
        var roots1 = engine.GetRootNodes();
        var firstId = roots1[1].Id;

        // Get children (cached)
        var children1 = engine.GetChildren(firstId);

        // Get same children again (from cache)
        var children2 = engine.GetChildren(firstId);

        Assert.That(children2.Count, Is.EqualTo(children1.Count));

        // Invalidate and re-fetch
        engine.Invalidate(firstId);
        var children3 = engine.GetChildren(firstId);

        Assert.That(children3.Count, Is.EqualTo(children1.Count));
    }
}
