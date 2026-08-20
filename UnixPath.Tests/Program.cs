using AIOrchestrator;
using AIOrchestrator.API;

namespace UnixPathTests;

// ═══════════════════════════════════════════════════════════════════
//  Unix path refactor — fast deterministic regression cycle.
//  Covers: UnixPath helpers, DocumentTool.ResolveSandboxPath round-trips
//  (OpenOrCreate/copyTo/CreateFromMarkdown/AddImage/ConvertToPdf/
//  Restore/ToMarkdown), sandbox boundary (../, drive paths, sibling
//  prefix), backward-compat relative names, result-message hygiene.
//  Direct method calls only — no LLM, runs in seconds.
// ═══════════════════════════════════════════════════════════════════
class Program
{
    static int _ok = 0, _fail = 0;

    static void Check(string name, bool cond, string? detail = null)
    {
        if (cond) { _ok++; Console.WriteLine($"  PASS  {name}"); }
        else { _fail++; Console.WriteLine($"  FAIL  {name}  {detail ?? ""}"); }
    }

    static bool ThrowsUnauthorized(Action a)
    {
        try { a(); return false; }
        catch (UnauthorizedAccessException) { return true; }
    }

    static void Main()
    {
        var testDir = Path.Combine(Path.GetTempPath(), "UnixPathTest_" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(testDir);
        Directory.CreateDirectory(Path.Combine(testDir, "sub"));
        Directory.CreateDirectory(Path.Combine(testDir, "md"));
        Setup.DocumentsPath = testDir;
        var root = Path.GetFullPath(testDir);

        // 1×1 PNG so AddImage has real image data to embed.
        File.WriteAllBytes(Path.Combine(testDir, "sub", "logo.png"), Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
        File.WriteAllText(Path.Combine(testDir, "md", "source.md"), "# From File\n\nMarkdown from a file path.");

        Console.WriteLine($"Unix path refactor regression | sandbox={root}\n");

        // ── A. UnixPath helpers ──
        Console.WriteLine("[A] UnixPath helpers");
        Check("ToAgent native → /folder/file.docx", UnixPath.ToAgent(@"folder\file.docx") == "/folder/file.docx");
        Check("ToAgent idempotent", UnixPath.ToAgent("/folder/file.docx") == "/folder/file.docx");
        Check("ToAgent strips drive", UnixPath.ToAgent(@"C:\docs\file.docx") == "/docs/file.docx");
        Check("ToAgent empty → empty", UnixPath.ToAgent("") == "");
        Check("ToNative /folder/file.docx → folder\\file.docx",
            UnixPath.ToNative("/folder/file.docx") == "folder" + Path.DirectorySeparatorChar + "file.docx");
        Check("ToNative plain name → plain name", UnixPath.ToNative("file.docx") == "file.docx");
        Check("Normalize all forms agree",
            UnixPath.Normalize("/folder/file.docx") == UnixPath.Normalize(@"folder\file.docx")
            && UnixPath.Normalize(@"folder\file.docx") == "folder/file.docx"
            && UnixPath.Normalize(@"C:/folder/file.docx") == "folder/file.docx");
        Check("Normalize strips trailing slash", UnixPath.Normalize("/folder/") == "folder");
        Check("GetDirectory /folder/file.docx → folder", UnixPath.GetDirectory("/folder/file.docx") == "folder");
        Check("GetDirectory root file → empty", UnixPath.GetDirectory("file.docx") == "");
        Check("GetDirectory backslash input", UnixPath.GetDirectory(@"folder\sub\f.docx") == "folder/sub");

        // ── B. DocumentTool Unix round-trips ──
        Console.WriteLine("[B] DocumentTool Unix round-trips");
        string msg;

        using (var w = new DocumentTool())   // one instance per file: a second instance cannot
        {                                // reopen a file this one still holds (FileShare.Read)
            msg = w.OpenOrCreate("/sub/a.docx");
            Check("OpenOrCreate creates /sub/a.docx", msg.StartsWith("Created 'a.docx'.") && File.Exists(Path.Combine(root, "sub", "a.docx")), msg);

            msg = w.AddParagraph("Hello Unix");
            Check("AddParagraph on Unix-opened doc", msg.StartsWith("Paragraph appended"), msg);

            msg = w.AddImage("/sub/logo.png");
            Check("AddImage with Unix imagePath", msg.StartsWith("Image inserted"), msg);

            msg = w.ToMarkdown();
            Check("ToMarkdown contains paragraph", msg.Contains("Hello Unix"), msg);

            msg = w.OpenOrCreate("/sub/a.docx");
            Check("OpenOrCreate reopens existing", msg.StartsWith("Opened 'a.docx'."), msg);

            msg = w.OpenOrCreate("/sub/a.docx", copyTo: "/sub/a_v2.docx");
            Check("copyTo Unix → copy created", msg.StartsWith("Opened 'a.docx' as a copy") && File.Exists(Path.Combine(root, "sub", "a_v2.docx")), msg);
            w.AddParagraph("Copied edit");
        }

        using (var w2 = new DocumentTool())
        {
            w2.OpenOrCreate("/sub/a.docx");
            var orig = w2.ToMarkdown();
            Check("original untouched by copyTo", orig.Contains("Hello Unix") && !orig.Contains("Copied edit"), orig);
        }

        using (var w3 = new DocumentTool())
        {
            msg = w3.OpenOrCreate("/sub/a_v2.docx");
            Check("copy opens and contains its own edit", msg.StartsWith("Opened 'a_v2.docx'") && w3.ToMarkdown().Contains("Copied edit"), msg);

            // Rollback: the tool restores the OPEN document (closes its handle, GitSupport.Restore, reopens).
            w3.OpenOrCreate("/sub/a.docx");
            w3.AddParagraph("To be reverted");
            Check("pre-restore edit persisted", w3.ToMarkdown().Contains("To be reverted"));
            var history = AIOrchestrator.GitSupport.History(AIOrchestrator.SandboxPath.Resolve("/sub/a.docx"));
            var restored = w3.Restore(history[^1].VersionId);   // oldest version
            Check("Restore succeeds", restored.StartsWith("Restored"), restored);
            var reverted = w3.ToMarkdown();
            Check("Restore reverts edit", reverted.Contains("Hello Unix") && !reverted.Contains("To be reverted"), reverted);

            // CreateFromMarkdown: inline content (bold stays as markdown in ToMarkdown output)
            msg = w3.CreateFromMarkdown("/md/plan.docx", markdown: "# Plan\n\nHello **world**");
            Check("CreateFromMarkdown inline → created", msg.StartsWith("Created 'plan.docx'") && File.Exists(Path.Combine(root, "md", "plan.docx")), msg);
            var planMd = w3.ToMarkdown();
            Check("CreateFromMarkdown content correct", planMd.Contains("Plan") && planMd.Contains("Hello"), planMd);

            // CreateFromMarkdown: from Unix markdownFile
            msg = w3.CreateFromMarkdown("/md/plan2.docx", markdownFile: "/md/source.md");
            Check("CreateFromMarkdown markdownFile → created", msg.StartsWith("Created 'plan2.docx'") && File.Exists(Path.Combine(root, "md", "plan2.docx")), msg);
            Check("markdownFile content loaded", w3.ToMarkdown().Contains("From File"), w3.ToMarkdown());

            msg = w3.CreateFromMarkdown("/md/x.docx", markdownFile: "/md/missing.md");
            Check("markdownFile missing → clear error", msg.Contains("Markdown file '/md/missing.md' not found"), msg);

            // ConvertToPdf: valid path reaches the engine check; missing path is caught first.
            msg = w3.ConvertToPdf("/sub/a.docx");
            Check("ConvertToPdf valid Unix path → engine error (not 'not found')",
                msg.Contains("requires a real PDF engine") && !msg.Contains("not found"), msg);
            msg = w3.ConvertToPdf("/nope.docx");
            Check("ConvertToPdf missing → not found", msg.Contains("File '/nope.docx' not found"), msg);

            // Backward compat: plain relative name still resolves inside the sandbox.
            msg = w3.OpenOrCreate("plain.docx");
            Check("backward-compat relative name", msg.StartsWith("Created 'plain.docx'.") && File.Exists(Path.Combine(root, "plain.docx")), msg);
        }

        // ── C. Sandbox boundary ──
        Console.WriteLine("[C] Sandbox boundary");
        using (var w = new DocumentTool())
        {
            w.OpenOrCreate("/sub/a.docx");   // AddImage needs an open document
            msg = w.AddImage("/../evil.png");
            Check("rejects AddImage traversal", msg.Contains("escapes the workspace sandbox"), msg);

            msg = w.OpenOrCreate("/../evil.docx");
            Check("rejects /../ traversal", msg.Contains("escapes the workspace sandbox"), msg);

            msg = w.OpenOrCreate(@"C:\Windows\evil.docx");
            Check("rejects absolute drive path outside", msg.Contains("escapes the workspace sandbox"), msg);

            msg = w.OpenOrCreate(root + "_evil\\file.docx");
            Check("rejects sibling-prefix path", msg.Contains("escapes the workspace sandbox"), msg);

            msg = w.OpenOrCreate("/sub/a.docx", copyTo: "/../evil.docx");
            Check("rejects copyTo traversal", msg.Contains("escapes the workspace sandbox"), msg);

            msg = w.CreateFromMarkdown("/../evil.docx", markdown: "x");
            Check("rejects CreateFromMarkdown traversal", msg.Contains("escapes the workspace sandbox"), msg);
        }

        // ── C2. SandboxPath helper ──
        Console.WriteLine("[C2] SandboxPath helper");
        Check("Resolve /sub/a.docx → inside", SandboxPath.Resolve("/sub/a.docx") == Path.Combine(root, "sub", "a.docx"));
        Check("Resolve plain name → inside", SandboxPath.Resolve("plain.docx") == Path.Combine(root, "plain.docx"));
        Check("Resolve rejects /../ traversal", ThrowsUnauthorized(() => SandboxPath.Resolve("/../evil.docx")));
        Check("Resolve rejects sibling-prefix path", ThrowsUnauthorized(() => SandboxPath.Resolve(root + "_evil\\file.docx")));
        Check("Resolve rejects absolute drive outside", ThrowsUnauthorized(() => SandboxPath.Resolve(@"C:\Windows\evil.docx")));
        Check("TryResolve false on escape", !SandboxPath.TryResolve("/../evil.docx", out _));
        Check("TryResolve true inside", SandboxPath.TryResolve("/sub/logo.png", out var rp) && rp == Path.Combine(root, "sub", "logo.png"));
        Check("ToAgent host inside → /sub/a.docx", SandboxPath.ToAgent(Path.Combine(root, "sub", "a.docx")) == "/sub/a.docx");
        Check("ToAgent host root → /", SandboxPath.ToAgent(root) == "/", $"got '{SandboxPath.ToAgent(root)}'");
        Check("ToAgent host outside → unchanged", SandboxPath.ToAgent(Path.Combine(root + "_evil", "x.docx")) == Path.Combine(root + "_evil", "x.docx"));
        Check("ToAgent empty → empty", SandboxPath.ToAgent("") == "");

        // ── D. Result-message hygiene (no native path / backslashes leak) ──
        Console.WriteLine("[D] Result-message hygiene");
        using (var w = new DocumentTool())
        {
            var results = new[]
            {
                w.OpenOrCreate("/sub/a.docx"),
                w.AddParagraph("hygiene"),
                w.AddImage("/sub/logo.png"),
                w.CreateFromMarkdown("/md/hygiene.docx", markdown: "# H"),
                w.CreateFromMarkdown("/md/h2.docx", markdownFile: "/md/source.md"),
                w.ConvertToPdf("/sub/a.docx"),
                w.ConvertToPdf("/nope.docx"),
                w.ToMarkdown(),
            };
            bool clean = true;
            foreach (var r in results)
                if (r.Contains(root) || r.Contains('\\')) { clean = false; Console.WriteLine($"    leak: {r}"); }
            Check("no native path / backslash in results", clean);
        }

        // ── E. Singleton RagDocumentProcessor (Setup) + FileTool via the shared instance ──
        // Regression for the refactor: the whole app must use ONE processor (Setup's),
        // re-pointed by ChangePath when DocumentsPath changes; FileTool must not own one.
        // Setup.RagDocumentProcessor is internal — verified here through its observable
        // effects (index marker + FileSearch target dir) rather than direct access.
        Console.WriteLine("[E] Singleton processor + ChangePath + FileTool");

        Setup.DocumentsPath = testDir; // same path → ChangePath no-op, no crash
        Setup.WaitForIndexIdle(TimeSpan.FromSeconds(30)); // settle the very first (async) build

        // Different sandbox → the shared processor re-indexes the new root (marker written).
        var dir2 = Path.Combine(Path.GetTempPath(), "UnixPathTest2_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(dir2, "docs"));
        File.WriteAllText(Path.Combine(dir2, "docs", "Contratto.md"), "# Contratto\n\nFirma Rossi");
        try
        {
            Setup.DocumentsPath = dir2;
            // The rebuild is async now — wait for it before checking the marker.
            Check("ChangePath re-indexed dir2 (marker written)",
                Setup.WaitForIndexIdle(TimeSpan.FromSeconds(30)) && File.Exists(Path.Combine(dir2, ".md", ".indexed-v2")));

            // FileTool now searches through the SHARED processor — no own instance, no own index.
            var ft = new FileTool();
            var res = ft.FileSearch(path: "/docs");
            Check("FileTool.FileSearch via shared processor finds /docs/Contratto",
                res.StartsWith('/') && res.Contains("Contratto"), res);

            // Re-pointing back to a previously indexed root must not re-index (marker present, Debug) — no crash.
            Setup.DocumentsPath = testDir;
            Setup.WaitForIndexIdle(TimeSpan.FromSeconds(30)); // let any in-flight build settle before cleanup
            Check("back to testDir → marker present, no re-index (no crash)",
                File.Exists(Path.Combine(root, ".md", ".indexed-v2")));
        }
        finally
        {
            try { Directory.Delete(dir2, true); } catch { }
        }

        Console.WriteLine($"\n═══ {_ok} passed, {_fail} failed ═══");
        try { AIOrchestrator.GitSupport.DeleteWorkspace(testDir); } catch { }
        Environment.ExitCode = _fail == 0 ? 0 : 1;
    }
}
