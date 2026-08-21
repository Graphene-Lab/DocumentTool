using System.Net.Http.Json;
using System.Text;
using System.Text.RegularExpressions;
using AIOrchestrator;
using AIOrchestrator.API;
using UISupportGeneric;

// ═══════════════════════════════════════════════════════════════════════════
//  CreateDocumentAgent.Tests — agent-driven document generation campaign
//
//  Verifies the long-running/completion-event architecture end to end:
//    • deterministic (no-LLM) tests: AgentTaskRegistry, AttachmentSandbox,
//      DocumentTool.CreateDocumentFromContext validation errors;
//    • agent tests (need the DeepSeekBridge at 127.0.0.1:8787): the agent must
//      discover the sandbox path of a chat attachment (the [Attachments
//      available locally] notice), pass it to create_document_from_context,
//      and — in responsive mode — receive the TaskCompleted event on the NEXT
//      ExecuteAction (AgentState.Initiative) and report the generated PDF.
//
//  GENERATION IS SIMULATED BY DEFAULT: DocumentTool.SimulateCreateDocument replaces
//  CreateDocument.Create (CreateInternal is already covered by its own tests;
//  a real run costs minutes + provider tokens) and reproduces the same artifacts
//  (PDF + ".info"). Pass --real to exercise the real pipeline (opt-in, slow).
//  The agent/LLM part is NOT simulated — the orchestrator, tools, registry and
//  event delivery run for real.
//
//  Build & run:   dotnet run --project CreateDocumentAgent.Tests -c Debug
//  Provider check: the bridge must respond on http://127.0.0.1:8787/v1/models
//  (see AGENT_TOOLS_GUIDE.md — Testing Agent Tools). LLM tests are skipped
//  with a clear message when the bridge is unreachable.
// ═══════════════════════════════════════════════════════════════════════════
namespace CreateDocumentAgentTests
{
    class Program
    {
        static int ok = 0, fail = 0, total = 0;
        static string sandbox = "";
        static readonly string ResultsFile = Path.Combine(Path.GetTempPath(), "createdoc_agent_test_results.txt");
        static void WriteResult(string line) => File.AppendAllText(ResultsFile, line + Environment.NewLine);

        static void Main(string[] args)
        {
            File.WriteAllText(ResultsFile, $"RUN {DateTime.Now:HH:mm:ss}\n");
            sandbox = Path.Combine(Path.GetTempPath(), "CreateDocAgent_" + Guid.NewGuid().ToString("N")[..6]);
            Directory.CreateDirectory(sandbox);
            Setup.DocumentsPath = sandbox; // sandbox root BEFORE any FileTool/RAG instantiation
            Setup.ProviderConfig = ProviderConfigs.Get("DeepSeekBridge");
            Log.IsEnabled = true;

            // "t5" runs only the synchronous flow (used to re-test after a transient bridge throttle
            // without re-running the ~20 min responsive generation).
            var onlyT5 = args.Contains("t5");
            // By default the heavy document generation (CreateDocument.Create → CreateInternal) is SIMULATED
            // via the DocumentTool.SimulateCreateDocument seam: CreateInternal is already covered by its own
            // tests, and a real run costs minutes + provider tokens. The simulation reproduces the same
            // observable artifacts (PDF + ".info"), so the orchestration checks stay meaningful. Pass --real
            // to exercise the actual pipeline end-to-end (opt-in, slow).
            var realGeneration = args.Contains("--real");
            if (!realGeneration)
                InstallDocumentSimulation();

            // Empirical measurement of the summarize compression ratio (--measure-summary <path.md>):
            // runs MarkdownSummarize.SummarizeFile via reflection (internal class) on a real file and prints
            // raw/summary sizes + ratio. Used to calibrate the sync pre-check of DocumentTool.CreateDocumentFromContext.
            var measureIdx = Array.IndexOf(args, "--measure-summary");
            if (measureIdx >= 0)
            {
                var sample = measureIdx + 1 < args.Length ? args[measureIdx + 1] : null;
                if (sample == null || !File.Exists(sample)) { Console.WriteLine("MEASURE: usage --measure-summary <path.md>"); return; }
                MeasureSummary(sample);
                return;
            }

            Console.WriteLine($"CreateDocumentAgent tests | sandbox: {sandbox} | generation: {(realGeneration ? "REAL" : "SIMULATED")}\n");
            WriteResult($"STARTED generation={(realGeneration ? "REAL" : "SIMULATED")}");

            // "direct" runs ONLY the fast deterministic tests (no LLM, no bridge, no agent) — used for a
            // quick smoke after changes.
            var onlyDirect = args.Contains("direct");
            if (onlyDirect)
            {
                TestRegistry();
                TestAttachmentSandbox();
                TestValidationErrors();
                TestNonMdContextResolution();
                TestCannotObtainMarkdown();
                TestWaitForCompletionTimeout();
                TestRegistryOrphanState();
                TestRegistryDisposeCleanup();
                TestDiscardOnDispose();
                TestSandboxEscape();
                TestSufficiencyCache();
                TestAttachmentSandboxNulls();
                Console.WriteLine($"\n{ok}/{total} passed, {fail} failed {(fail == 0 ? "ALL OK!" : "")}");
                WriteResult($"DONE {ok}/{total} passed, {fail} failed");
                return;
            }

            if (!onlyT5)
            {
                TestRegistry();
                TestAttachmentSandbox();
                TestValidationErrors();
            }

            // Direct (no-agent) tests: they call the target methods directly and check the expected result —
            // no orchestrator involved. T1-T10 are deterministic (no LLM); T12 needs the LLM (relevance
            // evaluation) and is guarded by the bridge check.
            if (!realGeneration)
                TestNonMdContextResolution(); // needs the simulation seam (with --real it would start a real generation)
            TestCannotObtainMarkdown();
            TestWaitForCompletionTimeout();
            TestRegistryOrphanState();
            TestRegistryDisposeCleanup();
            TestDiscardOnDispose();
            TestSandboxEscape();
            TestSufficiencyCache();
            TestAttachmentSandboxNulls();
            if (BridgeUp())
                TestFilterContextByRelevance(); // direct call, but needs the LLM (relevance evaluation)

            if (BridgeUp())
            {
                if (!onlyT5)
                    TestResponsiveAgentFlow();
                TestSyncAgentFlow();
            }
            else
            {
                Console.WriteLine("SKIP: DeepSeekBridge not reachable at http://127.0.0.1:8787 — agent tests skipped");
                WriteResult("SKIP agent-tests bridge-unreachable");
            }

            Console.WriteLine($"\n{ok}/{total} passed, {fail} failed {(fail == 0 ? "ALL OK!" : "")}");
            WriteResult($"DONE {ok}/{total} passed, {fail} failed");
        }

        static void Record(bool pass, string name, string detail)
        {
            total++;
            Console.WriteLine($"{(pass ? "PASS" : "FAIL")} {name}: {detail}");
            WriteResult($"{(pass ? "PASS" : "FAIL")} {name}: {detail}");
            if (pass) ok++; else fail++;
        }

        /// <summary>Installs the DocumentTool simulation seam: produces the same artifacts as the real pipeline
        /// (PDF + "# Cited documents:" .info in the output folder) and returns Success — without any LLM call.
        /// Mirrors CreateDocument.CreateInternal's custom-output behavior so outcome checks stay valid.</summary>
        static void InstallDocumentSimulation()
        {
            DocumentTool.SimulateCreateDocument = (dto, mdPaths, outDir) =>
            {
                Directory.CreateDirectory(outDir);
                var pdfName = CreateDocument.OutputFileName(dto.DocumentType?.ToUpper() ?? "DOC", dto.Subject, DateTime.UtcNow.ToString("yyyy-MM-dd"));
                var pdfPath = Path.Combine(outDir, pdfName);
                if (!File.Exists(pdfPath))
                    File.WriteAllBytes(pdfPath, [0x25, 0x50, 0x44, 0x46]); // "%PDF" stub
                var infoPath = Path.ChangeExtension(pdfPath, ".info");
                var lines = new List<string> { "# Cited documents:" };
                lines.AddRange((mdPaths ?? []).Select(ToSandboxForm));
                File.WriteAllLines(infoPath, lines);
                return DocumentCreationResult.Success;
            };
        }

        /// <summary>Host path → agent sandbox form when inside the workspace, else the file name
        /// (same rule as CreateDocument.ToSandboxPath).</summary>
        static string ToSandboxForm(string hostPath)
        {
            var root = Path.GetFullPath(sandbox);
            var prefix = root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return hostPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? UnixPath.ToAgent(Path.GetRelativePath(root, hostPath))
                : Path.GetFileName(hostPath);
        }

        // ── T1: AgentTaskRegistry round-trip (no LLM) ──
        static void TestRegistry()
        {
            try
            {
                var r = new AgentTaskRegistry();
                var text = r.Start("fake_method", () => "Success — done");
                var parsed = AgentTaskRegistry.TryParseTaskId($"ProcessingInProgress: 'fake_method' running, {AgentTaskRegistry.TaskIdMarker}{text}. An update will arrive when it completes.", out var taskId);
                var completed = r.WaitForCompletion(text, TimeSpan.FromSeconds(30));
                var consumed = r.TryGetCompletion(text, out var again);
                var ok1 = parsed && taskId == text;
                var ok2 = completed.Contains("TaskCompleted:") && completed.Contains(text.ToString()) && completed.Contains("Success — done");
                var ok3 = !consumed; // WaitForCompletion already consumed it
                Record(ok1 && ok2 && ok3, "T1 registry", $"parse={parsed} waitText='{completed[..Math.Min(completed.Length, 80)]}' consumedAgain={consumed}");
            }
            catch (Exception ex) { Record(false, "T1 registry", ex.Message); }
        }

        // ── T2: AttachmentSandbox persistence (no LLM) ──
        static void TestAttachmentSandbox()
        {
            try
            {
                var content = Encoding.UTF8.GetBytes("# Materiale interno\n\nProcedura di onboarding.\n");
                var attach = new FileAttachment("materiale_interno.md", content);
                var path1 = AttachmentSandbox.Persist(attach);
                var path2 = AttachmentSandbox.Persist(attach); // idempotent
                var ok1 = path1 != null && path1.StartsWith("/AIChatAttachments/") && path1 == path2;
                var originalHost = Path.Combine(sandbox, UnixPath.ToNative(path1!));
                var ok2 = File.Exists(originalHost);
                // shadow markdown with source: frontmatter pointing back to the original
                var mdDir = Path.Combine(sandbox, ".md");
                var shadow = Directory.Exists(mdDir) ? Directory.GetFiles(mdDir, "*.md").FirstOrDefault() : null;
                var ok3 = shadow != null && File.ReadAllText(shadow!).StartsWith("---\nsource: " + path1 + "\n---");
                Record(ok1 && ok2 && ok3, "T2 attachment sandbox", $"path='{path1}' originalExists={ok2} shadowOk={ok3}");
            }
            catch (Exception ex) { Record(false, "T2 attachment sandbox", ex.Message); }
        }

        // ── T3: DocumentTool.CreateDocumentFromContext validation errors (no LLM) ──
        static void TestValidationErrors()
        {
            try
            {
                var w = new DocumentTool();
                var r1 = w.CreateDocumentFromContext(new List<string> { "/x.md" }, "/out", "report", "s", format: "docx");
                var r2 = w.CreateDocumentFromContext(new List<string>(), "/out", "report", "s");
                var r3 = w.CreateDocumentFromContext(new List<string> { "/missing.md" }, "/out", "report", "s");
                var ok1 = r1.StartsWith("Error:");
                var ok2 = r2.StartsWith("Error:");
                var ok3 = r3.StartsWith("Error:");
                Record(ok1 && ok2 && ok3, "T3 validation", $"format={r1} empty={r2} missing={r3}");
            }
            catch (Exception ex) { Record(false, "T3 validation", ex.Message); }
        }

        // ── T6 (direct call, no agent): non-Markdown context file resolves to its /.md shadow ──
        // This is the user scenario: the agent passes an attachment path (non-md) to the tool, which must
        // resolve it to the shadow Markdown via the "source:" frontmatter. Needs the simulation seam
        // (skipped with --real, which would trigger a real generation).
        static void TestNonMdContextResolution()
        {
            try
            {
                var content = Encoding.UTF8.GetBytes("# Contratto di fornitura\n\nFornitore: ACME srl, importo 120.000 €, durata 24 mesi, penali 5%.\n");
                var attach = new FileAttachment("contratto_fornitura.txt", content); // .txt: converts to Markdown, NOT a .md passthrough
                var sandboxPath = AttachmentSandbox.Persist(attach);
                if (sandboxPath == null) { Record(false, "T6 non-md context resolution", "persist returned null"); return; }

                var w = new DocumentTool { AsyncTaskRegistry = new AgentTaskRegistry() };
                var result = w.CreateDocumentFromContext(new List<string> { sandboxPath }, "/report_txt", "report", "Contratto di fornitura");
                var taskId = ParseTaskId(result);
                if (taskId == null) { Record(false, "T6 non-md context resolution", $"no task id: {result}"); return; }
                var completed = w.AsyncTaskRegistry!.WaitForCompletion(taskId.Value, TimeSpan.FromSeconds(30));

                var outDir = Path.Combine(sandbox, "report_txt");
                var info = Directory.Exists(outDir) ? Directory.GetFiles(outDir, "*.info").FirstOrDefault() : null;
                var infoLines = info != null ? File.ReadAllLines(info) : [];
                // The cited document must be the /.md shadow, not the raw original
                var ok1 = completed.Contains("Success");
                var ok2 = infoLines.Length > 1 && infoLines.Skip(1).All(l => l.Contains("/.md/", StringComparison.Ordinal));
                Record(ok1 && ok2, "T6 non-md context resolution", $"success={ok1} citedShadow={ok2} cited='{string.Join("; ", infoLines.Skip(1))}'");
            }
            catch (Exception ex) { Record(false, "T6 non-md context resolution", ex.Message); }
        }

        // ── T7 (direct call, no agent): existing file WITHOUT a shadow → "cannot obtain Markdown" ──
        static void TestCannotObtainMarkdown()
        {
            try
            {
                File.WriteAllText(Path.Combine(sandbox, "raw_note.txt"), "nota senza shadow");
                var w = new DocumentTool();
                var result = w.CreateDocumentFromContext(new List<string> { "/raw_note.txt" }, "/out", "report", "s");
                var ok = result.StartsWith("Error: cannot obtain Markdown");
                Record(ok, "T7 cannot obtain markdown", $"result='{result}'");
            }
            catch (Exception ex) { Record(false, "T7 cannot obtain markdown", ex.Message); }
        }

        // ── T8 (direct call, no agent): WaitForCompletion timeout path ──
        static void TestWaitForCompletionTimeout()
        {
            try
            {
                var r = new AgentTaskRegistry();
                var id = r.Start("slow_task", () => { Thread.Sleep(5000); return "done"; });
                var text = r.WaitForCompletion(id, TimeSpan.FromMilliseconds(300));
                var ok = text.Contains("still running after the wait timeout");
                Record(ok, "T8 wait timeout", $"text='{text}'");
            }
            catch (Exception ex) { Record(false, "T8 wait timeout", ex.Message); }
        }

        // ── T9 (direct call, no agent): context-sufficiency cache round-trip ──
        static void TestSufficiencyCache()
        {
            try
            {
                var file = "test_sufficiency_cache_" + Guid.NewGuid().ToString("N")[..8] + ".bin";
                var dict = new Dictionary<ulong, ContextSufficiencyResult> { [42] = new(true, false, "ok") };
                Cache.SaveContextSufficiencyCache(dict, file);
                var loaded = Cache.LoadContextSufficiencyCache(file);
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, file));
                var ok = loaded.TryGetValue(42, out var r) && r.IsSufficient && !r.IsInternetRetrievable && r.Reasoning == "ok";
                Record(ok, "T9 sufficiency cache round-trip", $"loaded={loaded.Count}");
            }
            catch (Exception ex) { Record(false, "T9 sufficiency cache round-trip", ex.Message); }
        }

        // ── T10 (direct call, no agent): AttachmentSandbox null cases ──
        static void TestAttachmentSandboxNulls()
        {
            try
            {
                var empty = AttachmentSandbox.Persist(new FileAttachment("vuoto.md", Array.Empty<byte>()));
                Record(empty == null, "T10 attachment null cases", $"emptyContent={empty == null}");
            }
            catch (Exception ex) { Record(false, "T10 attachment null cases", ex.Message); }
        }

        // ── T16 (direct call, no agent): definitive cleanup — Dispose frees EVERYTHING (stored completions
        // and running tasks) and a pending WaitForCompletion returns immediately (resolved by Dispose),
        // never blocking on a task that belongs to a dead conversation. This is the per-orchestrator
        // lifecycle guarantee the 1h TTL used to provide. ──
        static void TestRegistryDisposeCleanup()
        {
            try
            {
                var r = new AgentTaskRegistry();
                var id = r.Start("slow", () => { Thread.Sleep(3000); return "done"; });
                var sw = System.Diagnostics.Stopwatch.StartNew();
                r.Dispose(); // simulates the orchestrator being disposed while its task is still running
                var text = r.WaitForCompletion(id, TimeSpan.FromMilliseconds(500)); // must return immediately
                sw.Stop();
                var ok1 = sw.ElapsedMilliseconds < 1500; // did NOT block for the task's 3s
                var ok2 = !r.IsCompleted(id) && !r.IsRunning(id); // everything freed
                Record(ok1 && ok2, "T16 registry dispose cleanup", $"returnedIn={sw.ElapsedMilliseconds}ms text='{Truncate(text, 70)}' freed={ok2}");
            }
            catch (Exception ex) { Record(false, "T16 registry dispose cleanup", ex.Message); }
        }

        // ── T15 (direct call, no agent): AgentTaskRegistry.Discard frees completions immediately — a
        // stored completion is removed right away, and a still-running task's completion is dropped when
        // it finishes (no 1h TTL wait). This is what an orchestrator Dispose triggers for its own tasks. ──
        static void TestDiscardOnDispose()
        {
            try
            {
                var r = new AgentTaskRegistry();
                // Case 1: a stored (never consumed) completion is removed immediately by Discard.
                var id1 = r.Start("discard1", () => "done");
                var deadline = DateTime.UtcNow.AddSeconds(10);
                while (!r.IsCompleted(id1) && DateTime.UtcNow < deadline) Thread.Sleep(50);
                var storedBefore = r.IsCompleted(id1);
                r.Discard(id1);
                var ok1 = storedBefore && !r.IsCompleted(id1);

                // Case 2: a still-running task is discarded → its completion must NOT be stored when done.
                var id2 = r.Start("discard2", () => { Thread.Sleep(800); return "done"; });
                r.Discard(id2); // while running
                var ok2 = !r.IsCompleted(id2); // right after the discard (still running)
                Thread.Sleep(1500); // let the task finish
                var ok3 = !r.IsCompleted(id2) && !r.IsRunning(id2);
                Record(ok1 && ok2 && ok3, "T15 discard on dispose", $"storedRemoved={ok1} runningMarked={ok2} droppedOnFinish={ok3}");
            }
            catch (Exception ex) { Record(false, "T15 discard on dispose", ex.Message); }
        }

        // ── T14 (direct call, no agent): sandbox boundary is enforced — path traversal and absolute paths
        // outside the workspace root must be rejected, never resolved or read. ──
        static void TestSandboxEscape()
        {
            try
            {
                var w = new DocumentTool();
                var r1 = w.CreateDocumentFromContext(new List<string> { "../../evil.md" }, "/out", "report", "s");
                var r2 = w.CreateDocumentFromContext(new List<string> { "C:\\Windows\\system32\\drivers\\etc\\hosts" }, "/out", "report", "s");
                var ok1 = r1.StartsWith("Error:") && (r1.Contains("escapes") || r1.Contains("outside"));
                var ok2 = r2.StartsWith("Error:") && (r2.Contains("escapes") || r2.Contains("outside"));
                Record(ok1 && ok2, "T14 sandbox escape", $"traversal='{r1}' absolute='{r2}'");
            }
            catch (Exception ex) { Record(false, "T14 sandbox escape", ex.Message); }
        }

        // ── T13 (direct call, no agent): a task whose completion was already consumed is detectable as
        // "lost" (neither running nor completed) — the exact state the orchestrator drain uses to drop a
        // pending id instead of letting it linger forever. ──
        static void TestRegistryOrphanState()
        {
            try
            {
                var r = new AgentTaskRegistry();
                var id = r.Start("orphan", () => "done");
                r.WaitForCompletion(id, TimeSpan.FromSeconds(30)); // consumes the completion
                var running = r.IsRunning(id);
                var completed = r.IsCompleted(id);
                Record(!running && !completed, "T13 registry orphan state", $"running={running} completed={completed} (both false = lost state detected)");
            }
            catch (Exception ex) { Record(false, "T13 registry orphan state", ex.Message); }
        }

        // ── T12 (direct call, needs LLM): FilterContextByRelevance keeps relevant files, drops irrelevant ──
        // Verifies the extracted reduction pipeline (the same relevance evaluation used for internet
        // candidates): a file on-topic stays, an off-topic file is excluded — this is how the agent's
        // context is reduced WITHOUT asking the agent to trim it.
        static void TestFilterContextByRelevance()
        {
            try
            {
                var relevantPath = Path.Combine(sandbox, "relevant.md");
                var irrelevantPath = Path.Combine(sandbox, "irrelevant.md");
                File.WriteAllText(relevantPath, "# Contratto di fornitura ACME\n\n## Condizioni\n\nFornitore ACME srl, importo 120.000 €, durata 24 mesi.\n## Clausole\n\nPenali 5%, garanzia 12 mesi, consegna entro 60 giorni.\n## Pagamenti\n\nPagamento a 60 giorni dalla fattura.\n");
                File.WriteAllText(irrelevantPath, "# Ricette di cucina\n\n## Pasta al pomodoro\n\nIngredienti: pomodori, basilico, olio.\n## Tiramisù\n\nMascarpone, caffè, savoiardi.\n");
                var llm = new LLMUtility("DeepSeekBridge");
                var reduced = DocumentsSupport.FilterContextByRelevance(llm, [relevantPath, irrelevantPath], "report", "Contratto di fornitura");
                var ok1 = reduced.Contains(relevantPath);
                var ok2 = !reduced.Contains(irrelevantPath);
                Record(ok1 && ok2, "T12 filter by relevance", $"keptRelevant={ok1} droppedIrrelevant={ok2} kept={reduced.Count}");
            }
            catch (Exception ex) { Record(false, "T12 filter by relevance", ex.Message); }
        }

        /// <summary>Runs MarkdownSummarize.SummarizeFile (internal class, hence reflection) on a real file and
        /// prints the compression ratio — useful to calibrate expectations about summarization (measured
        /// 0.7% on a single-chunk 175KB sample on 2026-08-10).</summary>
        static void MeasureSummary(string mdPath)
        {
            try
            {
                var llm = new LLMUtility("DeepSeekBridge");
                var type = typeof(CreateDocument).Assembly.GetType("AIOrchestrator.MarkdownSummarize");
                if (type == null) { Console.WriteLine("MEASURE: MarkdownSummarize type not found"); return; }
                var method = type.GetMethod("SummarizeFile", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                if (method == null) { Console.WriteLine("MEASURE: SummarizeFile not found"); return; }
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var summaryPath = (string?)method.Invoke(null, new object[] { mdPath, llm });
                sw.Stop();
                if (summaryPath == null || !File.Exists(summaryPath)) { Console.WriteLine("MEASURE: summarize FAILED (null)"); return; }
                var rawLen = new FileInfo(mdPath).Length;
                var sumLen = new FileInfo(summaryPath).Length;
                var ratioPct = rawLen > 0 ? 100.0 * sumLen / rawLen : 0;
                Console.WriteLine($"MEASURE: raw={rawLen:N0} chars | summary={sumLen:N0} chars | ratio={ratioPct:F1}% | elapsed={sw.Elapsed.TotalSeconds:F1}s");
                Console.WriteLine($"MEASURE: summary file: {summaryPath}");
            }
            catch (Exception ex) { Console.WriteLine($"MEASURE: error: {ex.Message}"); }
        }

        static bool BridgeUp()
        {
            try
            {
                using var hc = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                var resp = hc.GetAsync("http://127.0.0.1:8787/v1/models").GetAwaiter().GetResult();
                return resp.IsSuccessStatusCode;
            }
            catch { return false; }
        }

        // ── T4 (responsive mode): attachment → create_document_from_context → completion event ──
        static void TestResponsiveAgentFlow()
        {
            try
            {
                var orch = new AgentHarness("DeepSeekBridge") { AsyncTaskDeliveryEnabled = true };
                string? initiative = null;
                orch.AgentProgress += (_, e) => { if (e.State == AgentHarness.AgentState.Initiative) initiative = e.Message; };

                var attachContent = Encoding.UTF8.GetBytes(
                    "# Procedura di onboarding\n\n## Finalità\n\nDescrivere i passaggi per l'inserimento di un nuovo dipendente.\n\n## Passi\n\n1. Contratto firmato dal candidato.\n2. Creazione account aziendali e badge.\n3. Formazione obbligatoria (sicurezza + privacy).\n4. Assegnazione del mentor.\n5. Revisione a 30 giorni con il manager.\n\n## Responsabilità\n\n- HR: passi 1-2.\n- IT: passo 2.\n- Manager: passi 4-5.\n");

                var prompt = "Crea un documento PDF completo e ben strutturato. Usa create_document_from_context passando il file allegato (vedi l'avviso sugli allegati disponibili in locale), con destinationFolder '/report_output', documentType 'report', subject 'Procedura di onboarding'. Attendi l'esito.";
                var result1 = orch.ExecuteAction(prompt,
                    new[] { "FileTool", "DocumentTool" },
                    attachments: new[] { new FileAttachment("procedura_onboarding.md", attachContent) },
                    maxIterations: 40);
                var inProgress = result1.Message ?? "";
                var taskId = ParseTaskId(inProgress);
                var ok1 = taskId != null;
                Record(ok1, "T4 run1 in-progress", $"message='{Truncate(inProgress, 90)}' taskId={taskId}");

                // Wait for the background generation to finish (bounded by 25 min).
                // IsCompleted is a NON-consuming peek on the orchestrator's per-conversation registry:
                // consuming here (TryGetCompletion) would steal the completion from the drain, and the
                // initiative event of run 2 would never fire.
                var deadline = DateTime.UtcNow.AddMinutes(25);
                var finished = false;
                while (DateTime.UtcNow < deadline)
                {
                    if (taskId != null && orch.TaskRegistry.IsCompleted(taskId.Value)) { finished = true; break; }
                    Thread.Sleep(10_000);
                }
                Record(finished, "T4 background completion", finished ? "completed within 25 min" : "TIMEOUT waiting for background task");
                if (!finished) return;

                // Run 2: the drain delivers the completion as an initiative turn before the user prompt.
                var result2 = orch.ExecuteAction("ok", new[] { "FileTool", "DocumentTool" }, maxIterations: 20);
                var outDir = Path.Combine(sandbox, "report_output");
                var pdf = Directory.Exists(outDir) ? Directory.GetFiles(outDir, "*.pdf").FirstOrDefault() : null;
                var info = pdf != null ? Path.ChangeExtension(pdf, ".info") : null;
                var ok2 = pdf != null && File.Exists(pdf);
                var ok3 = info != null && File.Exists(info) && File.ReadAllLines(info).FirstOrDefault() == "# Cited documents:";
                var ok4 = initiative != null && (initiative.Contains("report") || initiative.Contains("PDF") || initiative.Contains("document"));
                Record(ok2 && ok3 && ok4, "T4 artifact + initiative", $"pdf={pdf != null} info={ok3} initiative='{Truncate(initiative ?? "(none)", 90)}'");
            }
            catch (Exception ex) { Record(false, "T4 responsive flow", ex.Message); }
        }

        // ── T5 (synchronous mode, stateless orchestrator): the outcome is delivered inline ──
        // A RICH attachment is used on purpose: the LLM sufficiency evaluator must approve the context
        // so the happy path (PDF generated) is exercised end-to-end in sync mode. (A minimal attachment
        // would correctly be judged insufficient → InsufficientSources, which validates the failure path
        // but not the success path.)
        static void TestSyncAgentFlow()
        {
            try
            {
                var orch = new AgentHarness("DeepSeekBridge"); // AsyncTaskDeliveryEnabled = false (default)
                var attachContent = Encoding.UTF8.GetBytes(
                    "# Progetto Aquila — Documento di progetto\n\n" +
                    "## 1. Contesto e motivazione\n\n" +
                    "Il portale clienti attuale è obsoleto, con tempi di risposta superiori a 8 secondi e un tasso di abbandono del 45% nelle operazioni di pagamento. Il progetto Aquila nasce per sostituirlo con una piattaforma moderna, accessibile e integrata con il CRM aziendale e i sistemi di fatturazione elettronica.\n\n" +
                    "## 2. Obiettivi (SMART)\n\n" +
                    "O1: Ridurre i tempi di risposta medi sotto i 2 secondi entro il Q3.\n" +
                    "O2: Portare l'abbandono delle operazioni di pagamento sotto il 15% entro il Q4.\n" +
                    "O3: Raggiungere un punteggio di soddisfazione utente (CSAT) ≥ 4,2/5 dopo il primo semestre di esercizio.\n" +
                    "O4: Copertura del 100% dei processi di onboarding clienti digitalizzati.\n\n" +
                    "## 3. Ambito e fuori ambito\n\n" +
                    "IN AMBITO: area autenticazione e profilo, gestione contratti, fatture e pagamenti, ticketing assistenza, dashboard cliente.\n" +
                    "FUORI AMBITO: app mobile nativa (solo web responsive), integrazione con l'ERP legacy (rimandata alla fase 3).\n\n" +
                    "## 4. Architettura di riferimento\n\n" +
                    "Frontend: SPA React 18 con TypeScript; Backend: API REST .NET 8; DB: PostgreSQL 16 con replica; Cache: Redis; Infrastruttura: Kubernetes on-premise; Observabilità: OpenTelemetry + Grafana.\n\n" +
                    "## 5. Fasi, durate e deliverable\n\n" +
                    "Fase 1 — Analisi e requisiti (4 settimane): deliverable = documento requisiti funzionali e non funzionali, modello dati v1.\n" +
                    "Fase 2 — Prototipo UI/UX (3 settimane): deliverable = prototipo cliccabile, test utente con 10 utenti, report usabilità.\n" +
                    "Fase 3 — Sviluppo backend (10 settimane): deliverable = API complete, integrazioni CRM/fatturazione, test automatici con copertura ≥ 80%.\n" +
                    "Fase 4 — Sviluppo frontend (8 settimane, in parallelo con parte della Fase 3): deliverable = SPA completa, responsive, accessibile WCAG 2.1 AA.\n" +
                    "Fase 5 — Collaudo e accettazione (4 settimane): deliverable = piano di test eseguito, report di collaudo, verbale di accettazione.\n" +
                    "Fase 6 — Rilascio e rodaggio (3 settimane): deliverable = go-live, piano di rollback, monitoraggio KPI, report di rodaggio.\n\n" +
                    "## 6. Organizzazione del team\n\n" +
                    "Project Manager: Mario Rossi. Analista funzionale: Elena Bianchi. Architetto: Luca Verdi. Sviluppatori backend: Anna Neri, Paolo Gialli. Sviluppatori frontend: Giulia Blu, Marco Arancio. QA: Federica Rosa. DevOps: Stefano Grigi. Owner di prodotto: Davide Viola.\n\n" +
                    "## 7. Rischi principali e mitigazioni\n\n" +
                    "R1 — Ritardo sull'integrazione CRM (probabilità alta, impatto alto): mitigazione = early spike in fase 1, contratto API congelato in fase 3.\n" +
                    "R2 — Cambio requisiti a metà sviluppo (probabilità media, impatto alto): mitigazione = backlog prioritizzato, sprint review settimanali, change control board.\n" +
                    "R3 — Disponibilità ridotta del team QA (probabilità media, impatto medio): mitigazione = automazione dei test, cross-training.\n\n" +
                    "## 8. KPI di successo e monitoraggio\n\n" +
                    "KPI1: tempo di risposta p95 < 2 s. KPI2: abbandono pagamenti < 15%. KPI3: CSAT ≥ 4,2/5. KPI4: uptime ≥ 99,9%. Dashboard settimanale condivisa con la direzione; revisione mensile dei KPI.\n\n" +
                    "## 9. Budget e vincoli\n\n" +
                    "Budget complessivo: 480.000 € (personale 350.000 €, infrastruttura 80.000 €, licenze 50.000 €). Vincolo principale: data di go-live inderogabile entro il 31 dicembre dell'anno in corso; il budget è bloccato, eventuali extra richiedono l'approvazione della direzione IT.\n\n" +
                    "## 10. Criteri di accettazione\n\n" +
                    "Superamento del collaudo con zero difetti critici e bloccanti; KPI di rodaggio conformi per 30 giorni consecutivi; documentazione operativa consegnata e validata dal service desk.\n");

                var prompt = "Crea un documento PDF completo e ben strutturato. Usa create_document_from_context passando il file allegato (vedi l'avviso), destinationFolder '/report_sync', documentType 'report', subject 'Progetto Aquila'. Attendi l'esito.";
                var result = orch.ExecuteAction(prompt,
                    new[] { "FileTool", "DocumentTool" },
                    attachments: new[] { new FileAttachment("progetto_aquila.md", attachContent) },
                    maxIterations: 40);
                var outDir = Path.Combine(sandbox, "report_sync");
                var pdf = Directory.Exists(outDir) ? Directory.GetFiles(outDir, "*.pdf").FirstOrDefault() : null;
                var info = pdf != null ? Path.ChangeExtension(pdf, ".info") : null;
                var ok1 = pdf != null && File.Exists(pdf);
                var ok2 = info != null && File.Exists(info) && File.ReadAllLines(info).FirstOrDefault() == "# Cited documents:";
                // The agent answers in the conversation language (Italian here), so success is "con successo",
                // not the English token "Success" — verify the artifact (pdf + .info) and that the message is
                // not an error, never a literal token.
                var msg = result.Message ?? "";
                var ok3 = !msg.StartsWith("Error", StringComparison.OrdinalIgnoreCase) && !msg.Contains("Error:");
                Record(ok1 && ok2 && ok3, "T5 sync flow", $"pdf={pdf != null} info={ok2} message='{Truncate(msg, 90)}'");
            }
            catch (Exception ex) { Record(false, "T5 sync flow", ex.Message); }
        }

        static long? ParseTaskId(string text)
        {
            // The canonical marker is the tool result's "task id N" (lowercase, parsed exactly by
            // AgentTaskRegistry.TryParseTaskId on the tool RESULT — never paraphrased).
            // Here we parse the AGENT'S natural-language message to the user, which may paraphrase
            // ("task ID 4", "Task ID 4", ...), so a case-insensitive regex is required.
            var m = Regex.Match(text, @"task\s+id\s+(\d+)", RegexOptions.IgnoreCase);
            return m.Success && long.TryParse(m.Groups[1].Value, out var id) ? id : null;
        }

        static string Truncate(string s, int n) => s.Length <= n ? s : s[..n] + "...";
    }
}
