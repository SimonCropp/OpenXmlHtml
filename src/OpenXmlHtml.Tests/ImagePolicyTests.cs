[TestFixture]
public class ImagePolicyTests
{
    [Test]
    public void Deny_RejectsAll()
    {
        var policy = ImagePolicy.Deny();
        Assert.That(policy.IsAllowed("https://example.com/img.png"), Is.False);
        Assert.That(policy.IsAllowed(@"C:\Images\photo.png"), Is.False);
    }

    [Test]
    public void AllowAll_AcceptsAll()
    {
        var policy = ImagePolicy.AllowAll();
        Assert.That(policy.IsAllowed("https://example.com/img.png"), Is.True);
        Assert.That(policy.IsAllowed(@"C:\Images\photo.png"), Is.True);
    }

    [Test]
    public void SafeDomains_ExactMatch()
    {
        var policy = ImagePolicy.SafeDomains("example.com");
        Assert.That(policy.IsAllowed("https://example.com/img.png"), Is.True);
        Assert.That(policy.IsAllowed("https://other.com/img.png"), Is.False);
    }

    [Test]
    public void SafeDomains_SubdomainMatch()
    {
        var policy = ImagePolicy.SafeDomains("example.com");
        Assert.That(policy.IsAllowed("https://images.example.com/img.png"), Is.True);
        Assert.That(policy.IsAllowed("https://cdn.images.example.com/img.png"), Is.True);
    }

    [Test]
    public void SafeDomains_RejectsPartialMatch()
    {
        var policy = ImagePolicy.SafeDomains("example.com");
        Assert.That(policy.IsAllowed("https://notexample.com/img.png"), Is.False);
        Assert.That(policy.IsAllowed("https://example.com.evil.com/img.png"), Is.False);
    }

    [Test]
    public void SafeDirectories_AllowsMatchingPath()
    {
        var policy = ImagePolicy.SafeDirectories(Path.GetTempPath());
        var testPath = Path.Combine(Path.GetTempPath(), "test.png");
        Assert.That(policy.IsAllowed(testPath), Is.True);
    }

    [Test]
    public void SafeDirectories_RejectsNonMatchingPath()
    {
        var safeDir = Path.Combine(Path.GetTempPath(), "safe_dir_test");
        var policy = ImagePolicy.SafeDirectories(safeDir);
        var testPath = Path.Combine(Path.GetTempPath(), "unsafe", "test.png");
        Assert.That(policy.IsAllowed(testPath), Is.False);
    }

    [Test]
    public void SafeDirectories_PathTraversalProtection()
    {
        var safeDir = Path.Combine(Path.GetTempPath(), "safe_dir_test");
        var policy = ImagePolicy.SafeDirectories(safeDir);
        var traversalPath = Path.Combine(safeDir, "..", "unsafe", "test.png");
        Assert.That(policy.IsAllowed(traversalPath), Is.False);
    }

    [Test]
    public void SafeDirectories_RejectsMalformedFileUri()
    {
        var policy = ImagePolicy.SafeDirectories(Path.GetTempPath());
        Assert.That(policy.IsAllowed("file:///%"), Is.False);
    }

    [Test]
    public void SafeDirectories_RejectsInvalidPath()
    {
        var policy = ImagePolicy.SafeDirectories(Path.GetTempPath());
        Assert.That(policy.IsAllowed("\0invalid"), Is.False);
    }

    [Test]
    public void Filter_CustomPredicate()
    {
        var policy = ImagePolicy.Filter(src => src.Contains("allowed"));
        Assert.That(policy.IsAllowed("https://allowed.example.com/img.png"), Is.True);
        Assert.That(policy.IsAllowed("https://denied.example.com/img.png"), Is.False);
    }
}
