using AIOrchestrator;
using AIOrchestrator.API;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordToolTests {
    // ═══════════════════════════════════════════════════════════════
    //  WordTool — Realistic Document Tests with Quality Analysis
    //  Cycle 2: adds HARD scenarios targeting previously-uncovered
    //  methods (charts, images, restore, find/replace, paragraph
    //  surgery, full table editing, global fonts, mega-document).
    // ═══════════════════════════════════════════════════════════════
    class Program {
        static int ok = 0, fail = 0, total = 0;
        static string testDir = "";
        static readonly string ResultsFile = Path.Combine(Path.GetTempPath(), "wordtool_test_results.txt");

        static void WriteResult(string line) => File.AppendAllText(ResultsFile, line + Environment.NewLine);

        static void Main() {
            File.WriteAllText(ResultsFile, $"RUN {DateTime.Now:HH:mm:ss}\n");
            testDir = Path.Combine(Path.GetTempPath(), "WordHard_" + Guid.NewGuid().ToString("N")[..6]);
            Directory.CreateDirectory(testDir);
            Setup.DocumentsPath = testDir;
            Log.IsEnabled = true;

            // API key injected via environment (never committed to source). Not required by
            // DeepseekBridge, but kept so the same harness can switch to key-based providers.
            Setup.GeminiApiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY") ?? "";
            Setup.ProviderConfig = ProviderConfigs.Get("DeepSeekBridge");

            // A minimal valid 1x1 PNG so AddImage has something to insert.
            File.WriteAllBytes(Path.Combine(testDir, "logo.png"), Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));

            Console.WriteLine($"Hard WordTool tests (DeepseekBridge) | Temp: {testDir}\n");
            WriteResult("STARTED");

            // ── Cycle 1 regression: original suite ──
            RunTest(1, "Annual Business Report",
                "Create 'Annual_Report_2025.docx' — a sophisticated annual business report. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown with a proper document structure:\n" +
                "# Annual Report 2025\n\n## Executive Summary\n\nThe company delivered strong results: revenue reached $48M, up 22% year-over-year, with EBITDA margin at 18%.\n\n## Financial Performance\n\n|Metric|2024|2025|Change|\n|---|---|---|---|\n|Revenue|$39M|$48M|+22%|\n|EBITDA|$6M|$8.6M|+43%|\n|Net Profit|$4M|$5.5M|+37%|\n\n## Market Position\n\nWe strengthened our leadership in the European market and expanded into Asia.\n\n## Outlook 2026\n\nDouble-digit growth expected across all segments.\n" +
                "TASK 2: Make the main title use 'Title' style.\n" +
                "TASK 3: Make 'Executive Summary' Heading1, and 'Financial Performance' + 'Market Position' + 'Outlook 2026' Heading2.\n" +
                "TASK 4: Add a Table of Contents right after the title with title 'Table of Contents' covering levels 1-2.\n" +
                "TASK 5: Add header 'CONFIDENTIAL — ANNUAL REPORT 2025'.\n" +
                "TASK 6: Add page numbers to the footer.\n" +
                "TASK 7: Set page size A4, portrait orientation, margins 1 inch.\n" +
                "TASK 8: Set document properties: Title='Annual Report 2025', Author='CFO Office', Subject='2025 Financial Results', Keywords='annual, report, 2025, financial'.\n",
                "Annual_Report_2025.docx",
                expected: new[] { "Annual Report 2025", "Executive Summary", "Financial Performance", "Market Position", "Outlook 2026", "$48M" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                checkPageNumbers: true);

            RunTest(2, "Legal Service Agreement",
                "Create 'Service_Agreement.docx' — a professional legal services agreement. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# SERVICE AGREEMENT\n\n## 1. Parties\n\nThis Agreement is made between Acme Corporation (\"Client\") and Beta Consulting Ltd (\"Consultant\").\n\n## 2. Scope of Services\n\nThe Consultant shall provide strategic advisory services as detailed in Exhibit A.\n\n## 3. Term and Termination\n\n3.1 This Agreement commences on 1 January 2026 for a term of 12 months.\n3.2 Either party may terminate with 30 days written notice.\n\n## 4. Fees and Payment\n\n|Service|Rate|Billing|\n|---|---|---|\n|Advisory|$250/hour|Monthly|\n|Strategy Review|$5,000|Per engagement|\n|Retainer|$10,000/month|Monthly|\n\n## 5. Confidentiality\n\nBoth parties agree to protect confidential information.\n\n## 6. Limitation of Liability\n\nLiability is limited to fees paid in the preceding 6 months.\n" +
                "TASK 2: Make the title 'SERVICE AGREEMENT' use 'Title' style.\n" +
                "TASK 3: Make sections 1-6 use Heading1 style.\n" +
                "TASK 4: Add a comment on the Fees section: 'Legal review required before signing'.\n" +
                "TASK 5: Create a versioned copy of the document: OpenOrCreate(filePath: 'Service_Agreement.docx', copyTo: 'Service_Agreement_v2.docx').\n" +
                "TASK 6: Add a new clause after section 6 to the COPY: '## 7. Governing Law\\n\\nThis Agreement is governed by the laws of England.'\n" +
                "TASK 7: Open the ORIGINAL 'Service_Agreement.docx' again (no copyTo) and verify clause 7 is NOT present.\n" +
                "TASK 8: Set page margins to 1 inch, page size A4.\n" +
                "TASK 9: Set document properties: Title='Service Agreement', Author='Legal Department'.\n",
                "Service_Agreement.docx",
                expected: new[] { "SERVICE AGREEMENT", "Parties", "Scope of Services", "Fees and Payment", "Confidentiality", "$250/hour" },
                requiredHeadings: new[] { "Title", "Heading1" },
                absentPhrases: new[] { "Governing Law", "governed by the laws of England" });

            RunTest(3, "Product Launch Plan",
                "Create 'Product_Launch_Plan.docx' — a complete go-to-market plan. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Product Launch Plan\n\n## Product Overview\n\nNova AI is a next-generation analytics platform for SMBs.\n\n## Target Market\n\n- SMBs with 10-200 employees\n- Marketing and sales teams\n- Data-driven decision makers\n\n## Launch Timeline\n\n|Phase|Date|Owner|\n|---|---|---|\n|Beta|2026-02-01|Engineering|\n|Soft Launch|2026-03-15|Marketing|\n|Full Launch|2026-04-01|All|\n\n## Go-To-Market Strategy\n\nFreemium model with enterprise upgrade path.\n\n## Success Metrics\n\n- 5,000 signups in first 90 days\n- 20% conversion to paid\n- 4.5 star app store rating\n" +
                "TASK 2: Apply 'Title' style to the main title.\n" +
                "TASK 3: Make 'Product Overview', 'Launch Timeline', 'Go-To-Market Strategy' Heading1 and the rest Heading2.\n" +
                "TASK 4: Add a bulleted list after 'Target Market' with: 'Competitor analysis complete', 'Pricing approved by board'.\n" +
                "TASK 5: Add a numbered list after 'Go-To-Market Strategy' with: 'Phase 1: Beta program', 'Phase 2: Content marketing', 'Phase 3: Paid acquisition'.\n" +
                "TASK 6: Add header 'NOVA AI — CONFIDENTIAL'.\n" +
                "TASK 7: Add page numbers.\n" +
                "TASK 8: Set A4 landscape orientation.\n" +
                "TASK 9: Set document properties: Title='Product Launch Plan', Author='Product Team'.\n",
                "Product_Launch_Plan.docx",
                expected: new[] { "Product Launch Plan", "Product Overview", "Target Market", "Launch Timeline", "Go-To-Market Strategy", "5,000 signups" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                checkLists: true, checkPageNumbers: true);

            RunTest(4, "Technical Architecture Document",
                "Create 'Architecture.docx' — a detailed technical architecture document. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# System Architecture v2\n\n## Overview\n\nThe platform follows a microservices architecture with event-driven communication.\n\n## Components\n\n|Component|Technology|Purpose|\n|---|---|---|\n|API Gateway|Kong|Routing & auth|\n|Core Services|.NET 10|Business logic|\n|Data Store|PostgreSQL|Primary storage|\n|Message Bus|Kafka|Event streaming|\n\n## Data Flow\n\n1. Client request → API Gateway\n2. Gateway authenticates → routes to service\n3. Service publishes events → Kafka\n4. Consumers process → update read models\n\n## Deployment\n\n- Kubernetes cluster, 3 environments\n- CI/CD via GitHub Actions\n- Canary releases\n\n## Security\n\nZero-trust network, mTLS everywhere, secrets via Vault.\n" +
                "TASK 2: Make title 'Title' style.\n" +
                "TASK 3: Make 'Components', 'Data Flow', 'Deployment' Heading1.\n" +
                "TASK 4: Add a new row to the components table: [\"Cache\",\"Redis\",\"Session & rate limits\"]\n" +
                "TASK 5: Add a bookmark named 'DataFlow' on the 'Data Flow' heading paragraph.\n" +
                "TASK 6: Add a cross-reference link to 'DataFlow' at the end of the document with text 'See Data Flow'.\n" +
                "TASK 7: Set the title color to dark blue #1F3864 and make it bold.\n" +
                "TASK 8: Add footer 'Page {number}'.\n" +
                "TASK 9: Set landscape A4.\n" +
                "TASK 10: Set document properties: Title='System Architecture v2', Author='Platform Team'.\n",
                "Architecture.docx",
                expected: new[] { "System Architecture v2", "Components", "Data Flow", "Deployment", "Kubernetes", "Redis" },
                requiredHeadings: new[] { "Title", "Heading1" },
                checkTables: true, checkPageNumbers: true);

            RunTest(5, "Q4 Board Presentation",
                "Create 'Board_Presentation.docx' — a structured board presentation. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Q4 Board Presentation\n\n## Performance Highlights\n\nStrong quarter: revenue $14.2M (+18% QoQ), 340 new customers, NRR 112%.\n\n## Financial Summary\n\n|KPI|Q3|Q4|\n|---|---|---|\n|Revenue|$12M|$14.2M|\n|Gross Margin|68%|71%|\n|Cash Burn|$1.2M|$0.8M|\n\n## Key Initiatives\n\n- Enterprise expansion\n- AI product launch\n- International hiring\n\n## Risks & Mitigations\n\n|Risk|Impact|Mitigation|\n|---|---|---|\n|Market slowdown|High|Diversify segments|\n|Key person|Medium|Succession plan|\n\n## Ask\n\nAdditional $5M growth capital.\n" +
                "TASK 2: Title style on main title.\n" +
                "TASK 3: 'Performance Highlights' Heading1, others Heading2.\n" +
                "TASK 4: Add a Table of Contents after the title with title 'Agenda' covering levels 1-2.\n" +
                "TASK 5: Add a comment on 'Additional $5M growth capital': 'Discuss valuation with board'.\n" +
                "TASK 6: Set header 'Q4 BOARD — INTERNAL'.\n" +
                "TASK 7: Add page numbers.\n" +
                "TASK 8: Set document properties: Title='Q4 Board Presentation', Author='CEO Office'.\n" +
                "TASK 9: Convert to markdown and report the output.\n",
                "Board_Presentation.docx",
                expected: new[] { "Q4 Board Presentation", "Performance Highlights", "Financial Summary", "Key Initiatives", "$14.2M", "112%" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                checkPageNumbers: true);

            // ── Cycle 2: HARD scenarios — previously-uncovered methods ──
            RunTest(6, "Newsletter (chart + image)",
                "Create 'Newsletter.docx' — a marketing newsletter with a chart and an image. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Q3 Newsletter\n\n## From the CEO\n\nWe hit 10,000 customers this quarter — a milestone worth celebrating.\n\n## Product Updates\n\n- New analytics dashboard\n- Mobile app 2.0\n- API rate limits raised\n\n## Upcoming Events\n\n|Event|Date|Location|\n|---|---|---|\n|User Conference|2026-09-15|Berlin|\n|Webinar: AI Roadmap|2026-10-02|Online|\n" +
                "TASK 2: Make the title use 'Title' style.\n" +
                "TASK 3: Make 'From the CEO' and 'Product Updates' Heading1, 'Upcoming Events' Heading2.\n" +
                "TASK 4: Add a native BAR chart titled 'Revenue by Quarter' with ONE series named 'Revenue (k$)' and values [120, 180, 240, 310], categories ['Q1','Q2','Q3','Q4']. Place it right after the 'Product Updates' section content.\n" +
                "TASK 5: Insert the image file 'logo.png' right after the title paragraph.\n" +
                "TASK 6: Add footer text 'Page {page}' — the footer must contain a REAL page-number field.\n" +
                "TASK 7: Add header 'NEWSLETTER — INTERNAL'.\n" +
                "TASK 8: Set page size A4 with margins 0.75 inches on ALL four sides.\n" +
                "TASK 9: Set document properties: Title='Q3 Newsletter', Author='Marketing Team'.\n",
                "Newsletter.docx",
                expected: new[] { "Q3 Newsletter", "From the CEO", "Product Updates", "Upcoming Events", "10,000 customers" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                checkTables: true, checkChart: true, checkImage: true, checkPageNumbers: true);

            RunTest(7, "Document surgery (move/copy/replace)",
                "Create 'Surgery.docx' — restructure an existing document. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Release Notes v2.0\n\n## Introduction\n\nWelcome to version 2.0 of the platform.\n\n## New Features\n\n- Real-time collaboration\n- Offline mode\n- Advanced search\n\n## Bug Fixes\n\nThe team fixed 34 bugs this release.\n\n## Known Issues\n\nSome legacy workflows may need reconfiguration.\n\n## Contact\n\nReach us at support@acme.io.\n" +
                "TASK 2: Move the paragraph containing 'Contact' to the position right after the title paragraph (index 1).\n" +
                "TASK 3: Copy the paragraph 'Welcome to version 2.0 of the platform.' to the very end of the document.\n" +
                "TASK 4: Replace ALL occurrences of 'acme.io' with 'globex.com' throughout the document (find/replace).\n" +
                "TASK 5: Insert a page break right before the 'Known Issues' heading.\n" +
                "TASK 6: Delete the 'Bug Fixes' section heading AND its body paragraph ('The team fixed 34 bugs...').\n" +
                "TASK 7: Set document properties: Title='Release Notes v2.0', Author='Platform Team'.\n",
                "Surgery.docx",
                expected: new[] { "Release Notes v2.0", "Introduction", "Known Issues", "globex.com", "Contact" },
                requiredHeadings: new[] { "Title", "Heading1" },
                absentPhrases: new[] { "acme.io", "34 bugs" });

            RunTest(8, "Table engineering (all ops)",
                "Create 'Inventory.docx' — build and manipulate tables with every operation. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Inventory Report\n\n## Warehouse A\n\n|SKU|Item|Qty|\n|---|---|---|\n|A-100|Laptop|42|\n|A-200|Monitor|77|\n\n## Warehouse B\n\n|SKU|Item|Qty|\n|---|---|---|\n|B-100|Keyboard|120|\n" +
                "TASK 2: Make 'Inventory Report' 'Title' style, 'Warehouse A' and 'Warehouse B' Heading1.\n" +
                "TASK 3: In the FIRST table (Warehouse A), add a row at the end: ['A-300','Dock','35'].\n" +
                "TASK 4: In the SECOND table (Warehouse B), add a new column at the end with values ['15','22'] (one per existing row).\n" +
                "TASK 5: In the SECOND table, delete the row that contains 'Keyboard'.\n" +
                "TASK 6: In the FIRST table, set the cell at row 1, column 1 (0-based) to 'Laptop Pro'.\n" +
                "TASK 7: Read the FIRST table back with GetTableData and include its row count in your final report message.\n" +
                "TASK 8: Set document properties: Title='Inventory Report', Author='Ops Team'.\n",
                "Inventory.docx",
                expected: new[] { "Inventory Report", "Warehouse A", "Warehouse B", "Laptop Pro", "A-300", "Dock" },
                requiredHeadings: new[] { "Title", "Heading1" },
                checkTables: true,
                absentPhrases: new[] { "Keyboard" });

            RunTest(9, "Transform + Restore roundtrip",
                "Create 'Transform.docx' — transform a document, then restore it. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Strategy Memo\n\n## Context\n\nWe are entering the APAC market in 2026.\n\n## Objectives\n\n1. Reach 50 partners by Q2\n2. Launch localized pricing\n3. Hire a regional team\n\n## Risks\n\nCurrency volatility and regulatory delays are the main risks.\n" +
                "TASK 2: Add a Table of Contents after the title with title 'Contents', levels 1-2.\n" +
                "TASK 3: Add header 'STRATEGY — CONFIDENTIAL' AND page numbers in the footer (a REAL page-number field).\n" +
                "TASK 4: Add a comment on the 'Risks' heading: 'Reviewed by legal — ok'.\n" +
                "TASK 5: Add a hyperlink with text 'Partner Portal' pointing to https://partners.example.com on the last paragraph.\n" +
                "TASK 6: Set document properties: Title='Strategy Memo', Author='Strategy Office'.\n" +
                "TASK 7: Convert the document to markdown and report the first 200 characters.\n" +
                "TASK 8: Restore the document to its pre-session state, then verify the hyperlink is GONE (restore reverts all edits).\n",
                "Transform.docx",
                expected: new[] { "Strategy Memo", "Context", "Objectives", "APAC market", "regulatory delays" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                absentPhrases: new[] { "Partner Portal", "partners.example.com" });

            RunTest(10, "Formatting deep-dive (global font)",
                "Create 'Proposal.docx' — heavy formatting with a global font change. Complete ALL tasks:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Business Proposal\n\n## Executive Summary\n\nA proposal for a 12-month partnership with a total value of $240k.\n\n## Services\n\n1. Quarterly strategy reviews\n2. Dedicated support engineer\n3. Executive workshops\n\n## Pricing\n\n|Item|Amount|\n|---|---|\n|Retainer|$15,000/mo|\n|Workshops|$60,000/yr|\n\n## Next Steps\n\nSign by end of Q3 to lock in pricing.\n" +
                "TASK 2: Apply 'Title' style to the main title, then make it BOLD with color #1F3864 and centered.\n" +
                "TASK 3: Make 'Executive Summary' and 'Services' Heading1, 'Pricing' and 'Next Steps' Heading2.\n" +
                "TASK 4: Set the font of the ENTIRE document to 'Calibri' at 11pt.\n" +
                "TASK 5: Add light gray shading (#F2F2F2) to the 'Executive Summary' paragraph and a single bottom border to the title paragraph.\n" +
                "TASK 6: Set 1.5 line spacing on the 'Services' section body.\n" +
                "TASK 7: Set document properties: Title='Business Proposal', Author='Sales Dept'.\n",
                "Proposal.docx",
                expected: new[] { "Business Proposal", "Executive Summary", "Services", "Pricing", "$240k", "$15,000/mo" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                checkTables: true, checkFontName: true);

            RunTest(11, "Enterprise mega-document",
                "Create 'Annual_Business_Plan_2026.docx' — a complete business plan exercising every WordTool feature. Complete ALL tasks IN ORDER:\n" +
                "TASK 1: Create from markdown:\n" +
                "# Annual Business Plan 2026\n\n## Executive Summary\n\nThe plan targets $12M ARR by year end through product-led growth.\n\n## Market Analysis\n\n|Segment|Size|Growth|\n|---|---|---|\n|SMB|$4B|+18%|\n|Mid-Market|$6B|+12%|\n|Enterprise|$10B|+9%|\n\n## Strategy\n\n- Expand self-serve onboarding\n- Launch enterprise tier in Q2\n- Double the partner program\n\n## Financial Plan\n\n|KPI|2025|2026 Target|\n|---|---|---|\n|ARR|$8M|$12M|\n|Gross Margin|70%|74%|\n\n## Risks\n\nCompetitive pressure and hiring constraints.\n" +
                "TASK 2: Apply 'Title' style to the main title.\n" +
                "TASK 3: Make 'Executive Summary' and 'Strategy' Heading1; 'Market Analysis', 'Financial Plan', 'Risks' Heading2.\n" +
                "TASK 4: Add a Table of Contents right after the title with title 'Contents', levels 1-2.\n" +
                "TASK 5: Add a native LINE chart 'ARR Trajectory' with ONE series 'ARR ($M)' values [4, 6, 8, 12] and categories ['2023','2024','2025','2026'], placed after the Financial Plan table.\n" +
                "TASK 6: In the Market Analysis table add a row: ['Channel Partners','$2B','+25%'].\n" +
                "TASK 7: Add a bulleted list after 'Strategy' with: 'Hire 40 engineers', 'Ship Q1 product milestone'.\n" +
                "TASK 8: Add a numbered list after 'Risks' with: 'Mitigation: expand EMEA sales', 'Mitigation: contractor bench'.\n" +
                "TASK 9: Add header 'BUSINESS PLAN 2026 — DRAFT' and footer with page numbers (a REAL page-number field) — do both in ONE SetHeaderFooter call.\n" +
                "TASK 10: Set page size A4, portrait, margins 1 inch.\n" +
                "TASK 11: Set document properties: Title='Annual Business Plan 2026', Author='CFO Office', Subject='2026 Plan', Keywords='plan, 2026, growth'.\n" +
                "TASK 12: Convert to markdown and report its length.\n",
                "Annual_Business_Plan_2026.docx",
                expected: new[] { "Annual Business Plan 2026", "Executive Summary", "Market Analysis", "Financial Plan", "Risks", "Channel Partners", "$12M" },
                requiredHeadings: new[] { "Title", "Heading1", "Heading2" },
                checkTables: true, checkLists: true, checkChart: true, checkPageNumbers: true,
                maxIterations: 150);

            Console.WriteLine($"\n{ok}/{total} passed, {fail} failed {(fail == 0 ? "ALL OK!" : "")}");
            WriteResult($"DONE {ok}/{total} passed, {fail} failed");
        }

        static void RunTest(int num, string name, string prompt, string expectedFile,
            string[] expected, string[] requiredHeadings,
            bool checkLists = false, bool checkTables = false,
            bool checkChart = false, bool checkImage = false, bool checkPageNumbers = false,
            bool checkFontName = false, string[]? absentPhrases = null,
            int maxIterations = 90) {
            total++;
            Console.Write($"T{num}: {name}... ");

            try {
                Log.LogStep($"=== Test {num}: {name} ===");
                var orch = new AIOrchestrator.AgentOrchestrator("DeepSeekBridge");

                using var done = new ManualResetEventSlim();
                AIOrchestrator.AgentOrchestrator.AgentProgressEventArgs? final = null;
                orch.AgentProgress += (_, e) => {
                    if (e.State is AIOrchestrator.AgentOrchestrator.AgentState.Completed
                        or AIOrchestrator.AgentOrchestrator.AgentState.Failed)
                    { final = e; done.Set(); }
                };

                var task = Task.Run(() => orch.ExecuteAction(prompt, new[] { typeof(WordTool) }, maxIterations: maxIterations));
                done.Wait(TimeSpan.FromSeconds(30));
                var agentResult = task.GetAwaiter().GetResult();

                WriteResult($"T{num} EVENT:{final?.State} ITERS:{final?.Iteration} MS:{final?.TotalElapsedMs}");

                if (!string.IsNullOrEmpty(agentResult.Error)) {
                    Console.WriteLine($"FAIL: agent error: {agentResult.Error}");
                    fail++; WriteResult($"T{num} FAIL agent-error");
                    return;
                }

                var report = AnalyzeDocument(expectedFile, expected, requiredHeadings,
                    checkLists, checkTables, checkChart, checkImage, checkPageNumbers, checkFontName, absentPhrases);
                var verdict = report.Total >= 90 ? "EXCELLENT" : report.Total >= 75 ? "GOOD" : report.Total >= 50 ? "POOR" : "FAIL";

                Console.WriteLine($"\n   Analysis: {report.Points}/{report.Max} ({report.Total}%) → {verdict}");
                foreach (var line in report.Details)
                    Console.WriteLine($"     • {line}");

                if (report.Total >= 75) {
                    Console.WriteLine($"   → PASS ({verdict})");
                    ok++;
                    WriteResult($"T{num} PASS {verdict} score:{report.Total}% iters:{agentResult.Iterations} ms:{agentResult.TotalElapsedMs}");
                } else {
                    Console.WriteLine($"   → FAIL ({verdict})");
                    fail++;
                    WriteResult($"T{num} FAIL {verdict} score:{report.Total}%");
                }
                Log.LogStep($"Test {num}: analysis score {report.Total}% ({verdict}), {agentResult.Iterations} iters");
            }
            catch (Exception ex) {
                Console.WriteLine($"FAIL: {ex.GetType().Name}: {ex.Message}");
                fail++;
                WriteResult($"T{num} CRASH {ex.GetType().Name}: {ex.Message}");
                Log.LogStep($"Test {num}: CRASH STACK: {ex}");
            }
        }

        class DocReport {
            public int Points = 0, Max = 0;
            public int Total => Max == 0 ? 0 : (int)Math.Round(100.0 * Points / Max);
            public List<string> Details = new();
        }

        static DocReport AnalyzeDocument(string fileName, string[] expected, string[] requiredHeadings,
            bool checkLists, bool checkTables, bool checkChart, bool checkImage, bool checkPageNumbers,
            bool checkFontName, string[]? absentPhrases) {
            var r = new DocReport();
            var filePath = Path.Combine(testDir, fileName);

            r.Max += 10;
            if (File.Exists(filePath)) { r.Points += 10; r.Details.Add("document created"); }
            else { r.Details.Add("DOCUMENT NOT FOUND"); return r; }

            r.Max += 6;
            bool hasHeaderFooter = false, hasCoreProps = false, hasNumbering = false;
            bool hasChart = false, hasImage = false, hasPageField = false, hasHyperlink = false, hasCalibri = false;
            try {
                using var fs = File.OpenRead(filePath);
                using var pkg = WordprocessingDocument.Open(fs, false);
                var mp = pkg.MainDocumentPart;
                hasHeaderFooter = (mp?.HeaderParts?.Count() ?? 0) > 0 || (mp?.FooterParts?.Count() ?? 0) > 0;
                hasCoreProps = !string.IsNullOrEmpty(pkg.PackageProperties?.Title)
                            || !string.IsNullOrEmpty(pkg.PackageProperties?.Creator);
                hasNumbering = mp?.Document?.Body?.Descendants<ParagraphProperties>()
                    .Any(pp => pp.NumberingProperties != null) == true;
                hasChart = (mp?.ChartParts?.Count() ?? 0) > 0;
                hasImage = (mp?.ImageParts?.Count() ?? 0) > 0;
                hasPageField = mp?.FooterParts?.Any(fp => fp.Footer?.Descendants<FieldCode>()
                    .Any(f => f.Text?.Contains("PAGE") == true) == true) == true;
                hasHyperlink = mp?.Document?.Descendants<Hyperlink>().Any() == true;
                hasCalibri = mp?.Document?.Body?.Descendants<Run>()
                    .Any(rn => rn.RunProperties?.RunFonts?.Ascii?.Value?.Equals("Calibri", StringComparison.OrdinalIgnoreCase) == true) == true;
            }
            catch { }
            if (hasHeaderFooter && hasCoreProps) { r.Points += 6; r.Details.Add("header/footer + core properties embedded"); }
            else if (hasHeaderFooter) { r.Points += 4; r.Details.Add("header/footer embedded (core props missing)"); }
            else if (hasCoreProps) { r.Points += 3; r.Details.Add("core properties set (header/footer missing)"); }
            else r.Details.Add("no header/footer parts and no core properties found");

            using var w = new WordTool();
            var open = w.OpenOrCreate(fileName);
            r.Max += 10;
            if (!open.StartsWith("Error")) { r.Points += 10; r.Details.Add("document opens correctly (valid DOCX)"); }
            else { r.Details.Add($"cannot open: {open}"); return r; }

            var text = w.ToMarkdown();
            // DocSharp escapes markdown specials inside table cells (e.g. "A-300" → "A\-300",
            // "|" → "\|"). Unescape before phrase matching so content checks see the real text.
            text = text.Replace("\\-", "-").Replace("\\|", "|").Replace("\\_", "_").Replace("\\*", "*");
            r.Max += expected.Length * 6;
            int found = 0;
            foreach (var phrase in expected) {
                if (text.Contains(phrase, StringComparison.OrdinalIgnoreCase)) { found++; }
                else r.Details.Add($"missing content: \"{phrase}\"");
            }
            r.Points += found * 6;
            r.Details.Add($"content coverage {found}/{expected.Length} expected phrases");

            r.Max += 5;
            bool hasMultipleSections = text.Split('\n').Count(l => l.Trim().Length > 0) >= 8;
            if (hasMultipleSections) { r.Points += 5; r.Details.Add("multi-section structure present"); }
            else r.Details.Add("document is too thin (few paragraphs)");

            if (absentPhrases != null) {
                r.Max += 6;
                bool absentOk = true;
                foreach (var ap in absentPhrases) {
                    if (text.Contains(ap, StringComparison.OrdinalIgnoreCase)) { absentOk = false; r.Details.Add($"forbidden content still present: \"{ap}\""); }
                }
                if (absentOk) { r.Points += 6; r.Details.Add("absent-content check passed"); }
            }

            var styles = new HashSet<string>(requiredHeadings, StringComparer.OrdinalIgnoreCase);
            var applied = new HashSet<string>();
            foreach (var p in w.GetParagraphs().Split("{\"index\"", StringSplitOptions.RemoveEmptyEntries)) {
                if (p.Contains("\"style\":\"Heading1\"")) applied.Add("Heading1");
                if (p.Contains("\"style\":\"Heading2\"")) applied.Add("Heading2");
                if (p.Contains("\"style\":\"Title\"")) applied.Add("Title");
                if (p.Contains("\"style\":\"Heading3\"")) applied.Add("Heading3");
            }
            r.Max += styles.Count * 5;
            foreach (var s in styles) {
                if (applied.Contains(s)) { r.Points += 5; r.Details.Add($"style '{s}' applied"); }
                else r.Details.Add($"style '{s}' NOT applied");
            }

            if (checkTables || text.Contains("|")) {
                r.Max += 8;
                var info = w.GetDocumentInfo();
                bool hasTable = info.Contains("\"tables\":") && !info.Contains("\"tables\":0");
                if (hasTable) { r.Points += 8; r.Details.Add("table(s) present and valid"); }
                else r.Details.Add("no tables found");
            }

            if (checkLists) {
                r.Max += 6;
                if (hasNumbering) { r.Points += 6; r.Details.Add("real lists present (numbering properties)"); }
                else r.Details.Add("lists missing (no numbering properties in document)");
            }

            if (checkChart) {
                r.Max += 8;
                if (hasChart) { r.Points += 8; r.Details.Add("native chart part embedded"); }
                else r.Details.Add("chart missing (no chart part)");
            }

            if (checkImage) {
                r.Max += 8;
                if (hasImage) { r.Points += 8; r.Details.Add("image part embedded"); }
                else r.Details.Add("image missing (no image part)");
            }

            if (checkPageNumbers) {
                r.Max += 8;
                if (hasPageField) { r.Points += 8; r.Details.Add("REAL page-number field in footer"); }
                else r.Details.Add("page-number field MISSING (literal text does not count)");
            }

            if (checkFontName) {
                r.Max += 6;
                if (hasCalibri) { r.Points += 6; r.Details.Add("global font 'Calibri' applied to runs"); }
                else r.Details.Add("Calibri font not found on any run");
            }

            return r;
        }
    }
}
