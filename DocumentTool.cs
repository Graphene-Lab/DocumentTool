using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Markdown;

namespace AIOrchestrator.API
{
    /// <summary>
    /// Document (DOCX) operations for agent use: open/create, paragraphs, tables, headers/footers, charts, images.
    /// </summary>
    public class DocumentTool : BaseAgentTool, IDisposable, IFileTool
    {
        /// <summary>This tool can start long-running background document generations whose completion is
        /// delivered back to the conversation as a standard completion event (see AGENT_TOOLS_GUIDE.md).</summary>
        public override bool SupportsAsyncTasks => true;

        /// <summary>Throwaway registry used when the tool is called OUTSIDE an orchestrator (AgentTaskRegistry.Current
        /// is null, e.g. in tests): assigned here so the caller can wait on the completion. In production the
        /// orchestrator's per-conversation registry comes from the ambient and this stays null.</summary>
        public AgentTaskRegistry? AsyncTaskRegistry { get; set; }

        private WordprocessingDocument? _document;
        private FileStream? _fileStream;   // owned separately: Open(stream) does NOT dispose the stream
        private string _filePath = string.Empty;
        private bool _fileCreatedThisSession;   // OpenOrCreate created an EMPTY placeholder this session

        private static readonly Random _idRandom = new();
        private static readonly object _idLock = new();
        private static readonly Dictionary<string, JustificationValues> Align = new(StringComparer.OrdinalIgnoreCase)
        {
            ["left"] = JustificationValues.Left, ["center"] = JustificationValues.Center,
            ["right"] = JustificationValues.Right, ["justified"] = JustificationValues.Both
        };
        private static readonly Dictionary<string, PageOrientationValues> Orient = new(StringComparer.OrdinalIgnoreCase)
        {
            ["portrait"] = PageOrientationValues.Portrait, ["landscape"] = PageOrientationValues.Landscape
        };
        private static readonly Dictionary<string, (uint W, uint H)> PageSizes = new(StringComparer.OrdinalIgnoreCase)
        {
            ["a4"] = (11906, 16838), ["letter"] = (12240, 15840), ["legal"] = (12240, 20160)
        };
        private const string WmlNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>Parameterless constructor for agent activation. Call OpenOrCreate or CreateFromMarkdown before using other methods.</summary>
        public DocumentTool() { }

        /// <summary>Opens an existing DOCX file for editing, or creates a new one if it doesn't exist.
        /// Every edit is saved to disk automatically — no Save call needed.</summary>
        /// <param name="filePath">Path relative to the workspace root (Unix style, e.g. "/folder/file.docx").</param>
        /// <param name="copyTo">Optional: instead of editing <paramref name="filePath"/> directly, copy it to
        /// this path (Unix style, e.g. "/folder/file_v2.docx") and edit the copy (the original is left untouched). Versioning, equivalent to Save As.</param>
        /// <returns>Descriptive result message.</returns>
        public string OpenOrCreate(string filePath, string? copyTo = null)
        {
            CloseDocument(saveFirst: false);
            _filePath = string.Empty;
            _fileCreatedThisSession = false;
            try
            {
                var resolved = SandboxPath.Resolve(filePath);
                if (copyTo != null)
                {
                    var copy = SandboxPath.Resolve(copyTo);
                    if (!File.Exists(resolved))
                        return $"Error: '{filePath}' not found. copyTo works on an existing file; use OpenOrCreate(path) without copyTo to create a new one.";
                    File.Copy(resolved, copy, overwrite: true);
                    _document = OpenEditable(copy);
                    _filePath = copy;
                    GitSupport.Snapshot(copy, "DocumentTool copy");
                    Log.LogStep($"DocumentTool.OpenOrCreate: opened copy of '{filePath}' at '{copy}'");
                    return $"Opened '{SandboxPath.ToAgent(resolved)}' as a copy at '{SandboxPath.ToAgent(copy)}'. Original untouched.";
                }

                bool created = false;
                if (!File.Exists(resolved))
                {
                    var doc = WordprocessingDocument.Create(resolved, WordprocessingDocumentType.Document);
                    doc.AddMainDocumentPart();
                    EnsureStylesPart(doc.MainDocumentPart!);
                    doc.MainDocumentPart!.Document = new Document(new Body());
                    doc.MainDocumentPart.Document.Save();
                    doc.Dispose();
                    created = true;
                }
                _document = OpenEditable(resolved);
                _filePath = resolved;
                _fileCreatedThisSession = created;
                Log.LogStep($"DocumentTool.OpenOrCreate: {(created ? "created" : "opened")} '{resolved}'");
                return created ? $"Created '{SandboxPath.ToAgent(resolved)}'." : $"Opened '{SandboxPath.ToAgent(resolved)}'.";
            }
            catch (Exception ex)
            {
                return $"Error: Cannot open '{filePath}'. {ex.Message}";
            }
        }

        /// <summary>Releases the underlying document resources. Called automatically by the orchestrator.
        /// Explicitly saves pending in-memory changes before closing — the OpenXml package's
        /// dispose-autosave can corrupt the file when a stream-backed document is closed.</summary>
        void IDisposable.Dispose() => CloseDocument(saveFirst: true);

        /// <summary>Reverts the OPEN document to a version from the workspace git repo (list them with
        /// GitTool.history). The current state is saved as a new version first (the rollback is
        /// reversible), then the file is overwritten and the document is reopened. Use this when the
        /// document is open in this tool; GitTool.restore handles files that are not open.</summary>
        /// <param name="versionId">Version to restore, from GitTool.history().</param>
        /// <returns>Descriptive result message.</returns>
        public string Restore(string versionId)
        {
            if (_document == null || string.IsNullOrEmpty(_filePath)) return "Error: No document open. Open it first.";
            try
            {
                CloseDocument(saveFirst: true);   // release the open handle so the file can be overwritten
                var message = GitSupport.Restore(versionId, _filePath);
                _document = OpenEditable(_filePath);
                return message;
            }
            catch (Exception ex)
            {
                try { if (_document == null) _document = OpenEditable(_filePath); } catch { }
                return $"Error: Restore failed. {ex.Message}";
            }
        }

        /// <summary>Closes the current document, disposing the package AND its backing stream.</summary>
        private void CloseDocument(bool saveFirst)
        {
            if (_document != null)
            {
                try { if (saveFirst) { _document.MainDocumentPart?.Document?.Save(); _document.Save(); } } catch { }   // best-effort flush
                try { _document.Dispose(); } catch { }
                _document = null;
            }
            try { _fileStream?.Dispose(); } catch { }   // must release even if package Dispose threw, or the file stays locked
            _fileStream = null;
        }

        /// <summary>Returns a JSON overview of the document: paragraphs, headings, tables, charts, sections.
        /// Call this FIRST to understand the document before editing. charts counts native charts
        /// (AddChart); each chart sits in its own paragraph — see GetParagraphs() for its index.</summary>
        /// <returns>JSON: {"paragraphs":N,"headings":[...],"tables":N,"charts":N,"sections":N,...}</returns>
        public string GetDocumentInfo() => Run("GetDocumentInfo", body =>
        {
            var paras = body.Elements<Paragraph>().ToList();
            var tables = body.Elements<Table>().ToList();
            var totalText = string.Concat(paras.SelectMany(p => p.InnerText));
            Log.LogStep($"DocumentTool.GetDocumentInfo: {paras.Count} paragraphs, {tables.Count} tables");
            return JsonSerializer.Serialize(new
            {
                paragraphs = paras.Count,
                headings = paras.Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("Heading") == true)
                    .Select(h => new { level = h.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "", text = Truncate(h.InnerText, 80) }).ToList(),
                tables = tables.Count,
                tableDetails = tables.Select((t, i) => new
                {
                    index = i,
                    rows = t.Elements<TableRow>().Count(),
                    cols = t.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 0
                }).ToList(),
                charts = body.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().Count(g => (g.Uri?.Value ?? "").Contains("/chart")),
                sections = body.Elements<SectionProperties>().Count(),
                estimatedWords = totalText.Split((char[])[' ', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries).Length,
                characters = totalText.Length
            });
        });

        /// <summary>Adds a paragraph with the specified text. Appends at the end by default,
        /// or inserts at a specific 0-based position when <paramref name="index"/> is provided.</summary>
        /// <param name="text">Paragraph text.</param>
        /// <param name="index">Optional 0-based position where to insert (from GetParagraphs()). Omit to append at the end.</param>
        /// <returns>Descriptive result message.</returns>
        public string AddParagraph(string text, int? index = null) => Run("AddParagraph", body =>
        {
            var para = new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            if (index == null)
            {
                AppendBody(body, para);
                return Done($"DocumentTool.AddParagraph: appended '{Truncate(text, 60)}'", $"Paragraph appended. ({text.Length} chars)");
            }
            var err = InsertIndex(body, index, out var idx);
            return err ?? Done($"DocumentTool.AddParagraph: inserted at index {index}, text='{Truncate(text, 60)}'", para, idx, $"Paragraph inserted at index {index}. ({text.Length} chars)");
        });

        /// <summary>Replaces all occurrences of oldText with newText in the document body.</summary>
        /// <param name="oldText">Text to find.</param>
        /// <param name="newText">Replacement text.</param>
        /// <returns>Number of replacements made.</returns>
        public string FindReplace(string oldText, string newText) => Run("FindReplace", body =>
        {
            if (string.IsNullOrEmpty(oldText)) return "Error: oldText cannot be empty.";
            int count = 0;
            // A phrase may span multiple runs (markdown bold splits text), so search per paragraph, not per Text node.
            foreach (var para in body.Descendants<Paragraph>())
            {
                var texts = para.Descendants<Text>().ToList();
                if (texts.Count == 0) continue;
                int n = 0;
                foreach (var t in texts)
                    if (t.Text.Contains(oldText, StringComparison.Ordinal))
                    {
                        n += (t.Text.Length - t.Text.Replace(oldText, "", StringComparison.Ordinal).Length) / oldText.Length;
                        t.Text = t.Text.Replace(oldText, newText);
                    }
                // Cross-run match: merge into the first run, drop the rest.
                var now = para.Descendants<Text>().ToList();
                var joined = string.Concat(now.Select(t => t.Text));
                if (now.Count > 1 && joined.Contains(oldText, StringComparison.Ordinal))
                {
                    var firstRun = now[0].Ancestors<Run>().FirstOrDefault();
                    if (firstRun != null)
                    {
                        n += (joined.Length - joined.Replace(oldText, "", StringComparison.Ordinal).Length) / oldText.Length;
                        firstRun.RemoveAllChildren<Text>();
                        firstRun.Append(new Text(joined.Replace(oldText, newText)) { Space = SpaceProcessingModeValues.Preserve });
                        foreach (var t in now.Skip(1)) t.Ancestors<Run>().FirstOrDefault()?.Remove();
                    }
                }
                count += n;
            }
            return Done($"DocumentTool.FindReplace: '{Truncate(oldText, 40)}' → '{Truncate(newText, 40)}' ({count} replacements)", $"{count} replacement(s) made.");
        });

        /// <summary>Sets font properties on all text in the document body. Properties are optional (null = unchanged).</summary>
        /// <param name="fontName">Font family name (e.g. "Arial", "Times New Roman"). Null to skip.</param>
        /// <param name="fontSize">Font size in half-points (e.g. 24 = 12pt). Null to skip.</param>
        /// <param name="bold">True=bold, false=normal. Null to skip.</param>
        /// <param name="italic">True=italic, false=normal. Null to skip.</param>
        /// <returns>Descriptive result message.</returns>
        public string SetDocumentFont(string? fontName = null, int? fontSize = null, bool? bold = null, bool? italic = null) => Run("SetDocumentFont", body =>
        {
            int changed = 0;
            foreach (var run in body.Descendants<Run>())
            {
                ApplyRunProps(run, fontName, fontSize, bold, italic, null, null);
                changed++;
            }
            return Done($"DocumentTool.SetDocumentFont: font={fontName ?? "(unchanged)"}, size={fontSize?.ToString() ?? "-"}, bold={bold}, italic={italic} ({changed} runs)", $"Font updated on {changed} run(s).");
        });

        /// <summary>Adds a table with the specified data. First row is treated as header.
        /// Appends at the end by default, or inserts before a specific paragraph when <paramref name="index"/> is provided.</summary>
        /// <param name="rows">2D string array: rows[row][col]. First row = header.</param>
        /// <param name="index">Optional 0-based paragraph position where to insert the table (from GetParagraphs()). Omit to append at the end.</param>
        /// <returns>Descriptive result message.</returns>
        public string AddTable(string[][] rows, int? index = null) => Run("AddTable", body =>
        {
            if (rows == null || rows.Length == 0 || rows[0].Length == 0)
                return "Error: Table must have at least one row with one column.";

            var table = new Table(
                new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                // tblGrid is required by the schema BEFORE the first w:tr.
                new TableGrid(Enumerable.Range(0, rows[0].Length).Select(_ => new GridColumn())));
            foreach (var rowData in rows)
            {
                var tr = new TableRow();
                foreach (var cell in rowData)
                    tr.Append(new TableCell(new Paragraph(new Run(new Text(cell ?? "")))));
                table.Append(tr);
            }
            if (index != null)
            {
                var err = InsertIndex(body, index, out var idx);
                if (err != null) return err;
                InsertAt(body, table, idx);
            }
            else AppendBody(body, table);
            return Done($"DocumentTool.AddTable: {rows.Length} rows, {rows[0].Length} cols at index {index?.ToString() ?? "end"}", $"Table added: {rows.Length} rows × {rows[0].Length} columns.");
        });

        /// <summary>Returns table data as a JSON array. First row is header.</summary>
        /// <param name="tableIndex">0-based table index (from GetDocumentInfo() or GetTableData()).</param>
        /// <returns>JSON string with table data, or error message.</returns>
        public string GetTableData(int tableIndex) => Run("GetTableData", body =>
        {
            var table = GetTableAt(body, tableIndex, out var err);
            if (err != null) return err;
            var result = table!.Elements<TableRow>()
                .Select(r => r.Elements<TableCell>().Select(c => c.InnerText.Trim()).ToList()).ToList();
            Log.LogStep($"DocumentTool.GetTableData: table {tableIndex}, {result.Count} rows");
            return JsonSerializer.Serialize(result);
        });

        /// <summary>Sets the document title (the first Title or Heading 1 paragraph, or creates one
        /// at the top) and applies the "Title" paragraph style. Skips a Heading 1 that is the title
        /// of a Table of Contents (immediately followed by a TOC field paragraph) — updating that
        /// one would rename the TOC, not the document.</summary>
        /// <param name="title">Title text.</param>
        /// <returns>Descriptive result message.</returns>
        public string SetTitle(string title) => Run("SetTitle", body =>
        {
            var existing = body.Elements<Paragraph>()
                .FirstOrDefault(p => (p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Title" ||
                                      p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Heading1") &&
                                     !(p.NextSibling() is Paragraph next &&
                                       next.Descendants<FieldCode>().Any(f => f.Text?.Contains("TOC") == true)));
            if (existing != null)
            {
                // Drop runs AND hyperlinks (w:hyperlink is a paragraph child, not a Run).
                existing.RemoveAllChildren<Run>();
                existing.RemoveAllChildren<Hyperlink>();
                existing.Append(new Run(new Text(title) { Space = SpaceProcessingModeValues.Preserve }));
                existing.ParagraphProperties ??= new ParagraphProperties();
                existing.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = "Title" };
                return Done($"DocumentTool.SetTitle: updated existing title to '{Truncate(title, 60)}'", $"Title updated to '{title}'.");
            }
            var titlePara = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Title" }),
                new Run(new Text(title) { Space = SpaceProcessingModeValues.Preserve }));
            var first = body.Elements<Paragraph>().FirstOrDefault();
            if (first != null) body.InsertBefore(titlePara, first); else AppendBody(body, titlePara);
            return Done($"DocumentTool.SetTitle: created title '{Truncate(title, 60)}'", $"Title set to '{title}'.");
        });

        /// <summary>Returns all paragraphs as JSON with index, text, style, and type.
        /// type="chart"/"image" marks paragraphs holding a chart or picture — find their index
        /// here, then remove them with ParagraphOp(index, "delete").</summary>
        /// <returns>JSON: [{"index":0,"text":"...","style":"Heading1","type":"text","words":2}, ...]</returns>
        public string GetParagraphs() => Run("GetParagraphs", body =>
        {
            var result = body.Elements<Paragraph>().Select((p, i) => new
            {
                index = i,
                text = Truncate(p.InnerText, 120),
                style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "(none)",
                type = ParagraphType(p),
                words = p.InnerText.Split((char[])[' ', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries).Length
            }).ToList();
            Log.LogStep($"DocumentTool.GetParagraphs: {result.Count} paragraphs");
            return JsonSerializer.Serialize(result);
        });

        /// <summary>Performs an operation on a paragraph: delete, copy (duplicate), move, replace its text,
        /// or remove its comments / bookmarks. "delete" also removes the whole paragraph's content,
        /// including charts (AddChart) and images (AddImage) — those live in a paragraph.</summary>
        /// <param name="index">0-based paragraph index (from GetParagraphs()).</param>
        /// <param name="action">"delete", "copy", "move", "replace", "delete_comment", or "delete_bookmark".</param>
        /// <param name="text">New text for "replace", or bookmark name for "delete_bookmark". Ignored otherwise.</param>
        /// <param name="toIndex">Target position for "move". Ignored otherwise.</param>
        /// <returns>Descriptive result message.</returns>
        public string ParagraphOp(int index, string action, string? text = null, int? toIndex = null) => Run("ParagraphOp", body =>
        {
            var para = GetParagraphAt(body, index, out var err);
            if (err != null) return err;

            switch (action.Trim().ToLowerInvariant())
            {
                case "delete":
                    var delText = Truncate(para!.InnerText, 40);
                    para.Remove();
                    return Done($"DocumentTool.ParagraphOp: deleted index {index} ('{delText}')", $"Paragraph {index} deleted ('{delText}').");

                case "copy":
                    var clone = (Paragraph)para!.CloneNode(true);
                    // Bookmarks are positional anchors; cloning duplicates their name (OOXML
                    // requires unique bookmark names) and would point cross-references at the
                    // wrong target. Strip them from the clone.
                    foreach (var bs in clone.Descendants<BookmarkStart>().ToList()) bs.Remove();
                    foreach (var be in clone.Descendants<BookmarkEnd>().ToList()) be.Remove();
                    var pos = Math.Min(index + 1, body.Elements<Paragraph>().Count());
                    InsertAt(body, clone, pos);
                    return Done($"DocumentTool.ParagraphOp: copied index {index} to position {pos}", $"Paragraph {index} duplicated.");

                case "replace":
                    if (text == null) return "Error: 'text' is required for 'replace'.";
                    para!.RemoveAllChildren<Run>();
                    para.RemoveAllChildren<Hyperlink>();   // w:hyperlink is a paragraph child, not a Run
                    para.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
                    return Done($"DocumentTool.ParagraphOp: replaced index {index}, text='{Truncate(text, 60)}'", $"Paragraph {index} replaced. ({text.Length} chars)");

                case "move":
                    if (toIndex == null) return "Error: 'toIndex' is required for 'move'.";
                    var all = body.Elements<Paragraph>().ToList();
                    if (toIndex < 0 || toIndex >= all.Count)
                        return $"Error: toIndex {toIndex} out of range. Document has {all.Count} paragraphs (0-{all.Count - 1}).";
                    para!.Remove();
                    InsertAt(body, para, Math.Min(toIndex.Value, body.Elements<Paragraph>().Count()));
                    return Done($"DocumentTool.ParagraphOp: moved from {index} to {toIndex}", $"Paragraph moved from position {index} to {toIndex}.");

                case "delete_comment":
                    var refs = para!.Descendants<CommentReference>().Select(c => c.Id?.Value).Where(id => id != null).ToList();
                    if (refs.Count == 0) return $"No comments on paragraph {index}.";
                    foreach (var run in para.Descendants<Run>().Where(r => r.Descendants<CommentReference>().Any()).ToList()) run.Remove();
                    var commentsPart = _document?.MainDocumentPart?.WordprocessingCommentsPart;
                    if (commentsPart?.Comments != null)
                    {
                        foreach (var id in refs)
                            commentsPart.Comments.Elements<Comment>().FirstOrDefault(c => c.Id?.Value == id)?.Remove();
                        commentsPart.Comments.Save();
                    }
                    return Done($"DocumentTool.ParagraphOp: removed {refs.Count} comment(s) from paragraph {index}", $"Removed {refs.Count} comment(s) from paragraph {index}.");

                case "delete_bookmark":
                    if (string.IsNullOrEmpty(text)) return "Error: 'text' (bookmark name) is required for 'delete_bookmark'.";
                    var starts = para!.Descendants<BookmarkStart>().Where(b => b.Name?.Value == text).ToList();
                    if (starts.Count == 0) return $"No bookmark '{text}' on paragraph {index}.";
                    foreach (var bs in starts)
                    {
                        var bid = bs.Id?.Value;
                        bs.Remove();
                        foreach (var be in para.Descendants<BookmarkEnd>().Where(e => e.Id?.Value == bid).ToList()) be.Remove();
                    }
                    return Done($"DocumentTool.ParagraphOp: removed bookmark '{text}' from paragraph {index}", $"Bookmark '{text}' removed from paragraph {index}.");

                default:
                    return $"Error: Unknown action '{action}'. Use 'delete', 'copy', 'move', 'replace', 'delete_comment', or 'delete_bookmark'.";
            }
        });

        /// <summary>Formats a paragraph: style, alignment, font, line spacing, shading and/or border.
        /// Pass only the properties you want to change; omitted ones are left untouched.</summary>
        /// <param name="index">0-based paragraph index (from GetParagraphs()).</param>
        /// <param name="style">Style ID: "Heading1"-"Heading9", "Title", "Normal". Null to skip.</param>
        /// <param name="alignment">"left", "center", "right", or "justified". Null to skip.</param>
        /// <param name="fontName">Font family (e.g. "Arial"). Null to skip.</param>
        /// <param name="fontSize">Font size in half-points (24 = 12pt). Null to skip.</param>
        /// <param name="bold">True=bold, false=normal. Null to skip.</param>
        /// <param name="italic">True=italic, false=normal. Null to skip.</param>
        /// <param name="underline">True=underline, false=none. Null to skip.</param>
        /// <param name="colorHex">Font/border color in "#RRGGBB" format. Null to skip.</param>
        /// <param name="lineSpacing">"single", "1.5", "double", or a number in points. Null to skip.</param>
        /// <param name="shading">Paragraph background in "#RRGGBB" format. Null to skip.</param>
        /// <param name="borderSide">"top", "bottom", "left", "right", or "all". Null to skip.</param>
        /// <param name="borderStyle">"single", "double", "dotted", "dashed" (default single).</param>
        /// <param name="borderSize">Border width in eighths of a point (4 = 0.5pt, 8 = 1pt).</param>
        /// <returns>Descriptive result message.</returns>
        public string FormatParagraph(int index,
            string? style = null, string? alignment = null,
            string? fontName = null, int? fontSize = null, bool? bold = null, bool? italic = null, bool? underline = null, string? colorHex = null,
            string? lineSpacing = null, string? shading = null,
            string? borderSide = null, string? borderStyle = null, int? borderSize = null) => Run("FormatParagraph", body =>
        {
            var para = GetParagraphAt(body, index, out var err);
            if (err != null) return err;
            para!.ParagraphProperties ??= new ParagraphProperties();
            var pp = para.ParagraphProperties;

            if (style != null) pp.ParagraphStyleId = new ParagraphStyleId { Val = style };
            if (alignment != null)
            {
                if (!Align.TryGetValue(alignment, out var jv))
                    return $"Error: Invalid alignment '{alignment}'. Use: left, center, right, or justified.";
                pp.Justification = new Justification { Val = jv };
            }
            if (fontName != null || fontSize != null || bold.HasValue || italic.HasValue || underline.HasValue || colorHex != null)
                foreach (var run in para.Descendants<Run>())
                    ApplyRunProps(run, fontName, fontSize, bold, italic, underline, colorHex);
            if (lineSpacing != null)
            {
                int lineVal = lineSpacing.Trim().ToLowerInvariant() switch
                {
                    "single" => 240,
                    "1.5" => 360,
                    "double" => 480,
                    _ => int.TryParse(lineSpacing, out var pts) && pts > 0 ? pts * 20 : 240
                };
                pp.SpacingBetweenLines ??= new SpacingBetweenLines();
                pp.SpacingBetweenLines.Line = lineVal.ToString();
                pp.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;
            }
            if (shading != null)
            {
                pp.Shading ??= new Shading();
                pp.Shading.Fill = shading.TrimStart('#');
                pp.Shading.Val = ShadingPatternValues.Clear;
            }
            if (borderSide != null)
            {
                var borderVal = (borderStyle ?? "single").Trim().ToLowerInvariant() switch
                {
                    "double" => BorderValues.Double, "dotted" => BorderValues.Dotted,
                    "dashed" => BorderValues.Dashed, _ => BorderValues.Single
                };
                var size = new UInt32Value((uint)(borderSize ?? 4));
                var color = (colorHex ?? "#000000").TrimStart('#');
                var borders = pp.ParagraphBorders ??= new ParagraphBorders();
                switch (borderSide.Trim().ToLowerInvariant())
                {
                    case "top": borders.TopBorder = new TopBorder { Val = borderVal, Color = color, Size = size }; break;
                    case "bottom": borders.BottomBorder = new BottomBorder { Val = borderVal, Color = color, Size = size }; break;
                    case "left": borders.LeftBorder = new LeftBorder { Val = borderVal, Color = color, Size = size }; break;
                    case "right": borders.RightBorder = new RightBorder { Val = borderVal, Color = color, Size = size }; break;
                    case "all":
                        borders.TopBorder = new TopBorder { Val = borderVal, Color = color, Size = size };
                        borders.BottomBorder = new BottomBorder { Val = borderVal, Color = color, Size = size };
                        borders.LeftBorder = new LeftBorder { Val = borderVal, Color = color, Size = size };
                        borders.RightBorder = new RightBorder { Val = borderVal, Color = color, Size = size };
                        break;
                    default:
                        return $"Error: Invalid borderSide '{borderSide}'. Use: top, bottom, left, right, or all.";
                }
            }
            return Done($"DocumentTool.FormatParagraph: index {index}, style={style ?? "-"}, align={alignment ?? "-"}, font={fontName ?? "-"}, size={fontSize?.ToString() ?? "-"}, bold={bold}, italic={italic}, underline={underline}, color={colorHex ?? "-"}, spacing={lineSpacing ?? "-"}, shading={shading ?? "-"}, border={borderSide ?? "-"}", $"Paragraph {index} formatted.");
        });

        /// <summary>Adds a list (bulleted or numbered). Each string in items becomes one item.
        /// Appends at the end by default, or inserts before a specific paragraph when <paramref name="index"/> is provided.
        /// Lists are properly defined in the document.</summary>
        /// <param name="items">List item texts.</param>
        /// <param name="type">Must be exactly "bulleted" (default) or "numbered". Do NOT use
        /// synonyms like "unordered", "bullet", "numeric" — they are rejected.</param>
        /// <param name="index">Optional 0-based paragraph position where to insert the list (from GetParagraphs()). Omit to append at the end.</param>
        /// <returns>Descriptive result message.</returns>
        public string AddList(string[] items, string type = "bulleted", int? index = null) => Run("AddList", body =>
        {
            if (items == null || items.Length == 0) return "Error: Items array is empty.";
            bool numbered = type.Trim().ToLowerInvariant() == "numbered";
            if (!numbered && !type.Trim().Equals("bulleted", StringComparison.OrdinalIgnoreCase))
                return $"Error: Unknown list type '{type}'. Use 'bulleted' or 'numbered'.";

            var err = InsertIndex(body, index, out var insertIdx);
            if (err != null) return err;
            EnsureNumberingPart(_document!.MainDocumentPart!);
            int numId = numbered ? 2 : 1;
            foreach (var item in items)
                InsertAt(body, new Paragraph(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numId })),
                    new Run(new Text(item) { Space = SpaceProcessingModeValues.Preserve })), insertIdx++);
            return Done($"DocumentTool.AddList: {items.Length} items ({(numbered ? "numbered" : "bulleted")}) at index {index?.ToString() ?? "end"}", $"{(numbered ? "Numbered" : "Bulleted")} list added: {items.Length} items.");
        });

        /// <summary>Adds a reference element to a paragraph: hyperlink, bookmark, or cross-reference to a bookmark.</summary>
        /// <param name="paragraphIndex">0-based paragraph index (from GetParagraphs()).</param>
        /// <param name="type">"hyperlink" (target = URL), "bookmark" (target = bookmark name), or
        /// "cross_reference" (target = existing bookmark name). Null: inferred — target starting
        /// with http/https becomes "hyperlink".</param>
        /// <param name="target">URL for hyperlink, or bookmark name.</param>
        /// <param name="text">Visible text. Defaults to target for hyperlinks.</param>
        /// <returns>Descriptive result message.</returns>
        public string AddLink(int paragraphIndex, string? type = null, string? target = null, string? text = null) => Run("AddLink", body =>
        {
            if (string.IsNullOrEmpty(target)) return "Error: Target cannot be empty.";
            type ??= target.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                  || target.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                  ? "hyperlink" : "bookmark";
            var para = GetParagraphAt(body, paragraphIndex, out var err);
            if (err != null) return err;

            string kind;
            switch (type.Trim().ToLowerInvariant())
            {
                case "hyperlink":
                    var relId = _document!.MainDocumentPart!.AddHyperlinkRelationship(new Uri(target), true).Id;
                    para!.RemoveAllChildren<Run>();
                    para.RemoveAllChildren<Hyperlink>();   // replace an existing link: w:hyperlink is a paragraph child, not a Run
                    para.Append(new Hyperlink(LinkRun(text ?? target)) { Id = relId });
                    kind = "Hyperlink";
                    Log.LogStep($"DocumentTool.AddLink: hyperlink at {paragraphIndex} → '{target}'");
                    break;
                case "cross_reference":
                    para!.RemoveAllChildren<Run>();
                    para.RemoveAllChildren<Hyperlink>();
                    para.Append(new Hyperlink(LinkRun(text ?? target)) { Anchor = target });
                    kind = "Cross-reference";
                    Log.LogStep($"DocumentTool.AddLink: cross-reference at {paragraphIndex} → #{target}");
                    break;
                case "bookmark":
                    var id = NewId();
                    para!.InsertBefore(new BookmarkStart { Name = target, Id = id }, para.GetFirstChild<Run>());
                    para.Append(new BookmarkEnd { Id = id });
                    kind = "Bookmark";
                    Log.LogStep($"DocumentTool.AddLink: bookmark '{target}' at {paragraphIndex}");
                    break;
                default:
                    return $"Error: Unknown link type '{type}'. Use 'hyperlink', 'bookmark', or 'cross_reference'.";
            }
            Persist();
            return $"{kind} added to paragraph {paragraphIndex}.";
        });

        /// <summary>Sets page properties for all sections: size, orientation, and/or margins.
        /// Pass only the values you want to change; omitted ones are left untouched.</summary>
        /// <param name="pageSize">"A4", "Letter", or "Legal". Null to skip.</param>
        /// <param name="orientation">"portrait" or "landscape". Null to skip.</param>
        /// <param name="top">Top margin in inches (e.g. 1, 0.75). Null to skip.</param>
        /// <param name="bottom">Bottom margin in inches. Null to skip.</param>
        /// <param name="left">Left margin in inches. Null to skip.</param>
        /// <param name="right">Right margin in inches. Null to skip.</param>
        /// <param name="margins">Convenience: sets ALL FOUR margins at once, in inches —
        /// "1", "0.75", "2.54cm", "25.4mm" or "72pt" are all accepted. Overridden by
        /// top/bottom/left/right when both are given. Null to skip.</param>
        /// <returns>Descriptive result message.</returns>
        public string SetPage(string? pageSize = null, string? orientation = null,
            double? top = null, double? bottom = null, double? left = null, double? right = null,
            string? margins = null) => Run("SetPage", body =>
        {
            var sections = EnsureSections(body);

            if (pageSize != null)
            {
                if (!PageSizes.TryGetValue(pageSize.Trim(), out var size))
                    return $"Error: Unknown page size '{pageSize}'. Use A4, Letter, or Legal.";
                foreach (var s in sections)
                {
                    var ps = s.GetFirstChild<PageSize>() ?? s.AppendChild(new PageSize());
                    // Keep dims consistent with the current orientation (landscape swaps W/H).
                    bool landscape = ps.Orient?.Value == PageOrientationValues.Landscape;
                    ps.Width = new UInt32Value(landscape ? size.H : size.W);
                    ps.Height = new UInt32Value(landscape ? size.W : size.H);
                }
            }

            if (orientation != null)
            {
                if (!Orient.TryGetValue(orientation, out var ov))
                    return $"Error: Invalid orientation '{orientation}'. Use 'portrait' or 'landscape'.";
                foreach (var s in sections)
                {
                    var ps = s.GetFirstChild<PageSize>() ?? s.AppendChild(new PageSize());
                    if (ps.Width == null || ps.Height == null)   // fresh PageSize → default A4 so the swap has dimensions
                    {
                        ps.Width = new UInt32Value(11906U);
                        ps.Height = new UInt32Value(16838U);
                    }
                    ps.Orient = ov;
                    bool needsSwap = ov == PageOrientationValues.Landscape ? ps.Width!.Value < ps.Height!.Value : ps.Width!.Value > ps.Height!.Value;
                    if (needsSwap)
                        (ps.Width, ps.Height) = (new UInt32Value(ps.Height.Value), new UInt32Value(ps.Width.Value));
                }
            }

            if (margins != null)
            {
                var all = ParseInches(margins);
                if (all == null)
                    return $"Error: Invalid margins '{margins}'. Use inches as a number ('1', '0.5') or with units ('2.54cm', '25.4mm', '72pt', '1in').";
                top ??= all; bottom ??= all; left ??= all; right ??= all;
            }

            if (top.HasValue || bottom.HasValue || left.HasValue || right.HasValue)
                foreach (var s in sections)
                {
                    var pm = s.GetFirstChild<PageMargin>() ?? s.AppendChild(new PageMargin());
                    if (top.HasValue) pm.SetAttribute(new OpenXmlAttribute("w", "top", WmlNs, ((int)(top.Value * 1440)).ToString()));
                    if (bottom.HasValue) pm.SetAttribute(new OpenXmlAttribute("w", "bottom", WmlNs, ((int)(bottom.Value * 1440)).ToString()));
                    if (left.HasValue) pm.SetAttribute(new OpenXmlAttribute("w", "left", WmlNs, ((int)(left.Value * 1440)).ToString()));
                    if (right.HasValue) pm.SetAttribute(new OpenXmlAttribute("w", "right", WmlNs, ((int)(right.Value * 1440)).ToString()));
                }

            return Done($"DocumentTool.SetPage: size={pageSize ?? "-"}, orient={orientation ?? "-"}, margins=({top}, {bottom}, {left}, {right})", "Page settings updated.");
        });

        /// <summary>Returns the text content of a specific paragraph.</summary>
        /// <param name="index">0-based paragraph index (from GetParagraphs()).</param>
        /// <returns>Text content of the paragraph.</returns>
        public string GetParagraphText(int index) => Run("GetParagraphText", body =>
        {
            var para = GetParagraphAt(body, index, out var err);
            return err != null ? err : para!.InnerText;
        });

        /// <summary>Sets a core document property (title, author, subject, keywords).
        /// ONE property per call.</summary>
        /// <param name="property">The property NAME (not its value), one of: "Title", "Author",
        /// "Subject", "Keywords". Example: set_document_property(property:'Title', value:'Q3 Report').</param>
        /// <param name="value">Property value.</param>
        /// <returns>Descriptive result message.</returns>
        public string SetDocumentProperty(string property, string value)
        {
            if (_document == null) return "Error: No document open.";
            try
            {
                var p = _document.PackageProperties;
                if (p == null) return "Error: Document properties not available.";
                switch (property.Trim().ToLowerInvariant())
                {
                    case "title": p.Title = value; break;
                    case "author": p.Creator = value; break;
                    case "subject": p.Subject = value; break;
                    case "keywords": p.Keywords = value; break;
                    default:
                        return $"Error: Unknown property '{property}'. 'property' must be the NAME (Title, Author, Subject, or Keywords), with the value in the 'value' parameter — one per call. Example: set_document_property(property='Title', value='My Doc').";
                }
                return Done($"DocumentTool.SetDocumentProperty: {property} = '{value}'", $"Document property '{property}' set.");
            }
            catch (Exception ex)
            {
                return $"Error: SetDocumentProperty failed. {ex.Message}";
            }
        }

        /// <summary>Sets header/footer text and/or page numbers. Page numbers come ONLY from
        /// pageNumbers:true (bare number) or a "{page}"/"{number}" placeholder inside the text
        /// (e.g. "Page {page} of 3"). Any other footer text is plain text — it does NOT create a
        /// page-number field. Pass only what you want to change; omitted ones are left untouched.
        /// Empty string ("") REMOVES the header/footer.</summary>
        /// <param name="header">Header text to set. "" removes the header. Null to skip.</param>
        /// <param name="footer">Footer text. Put "{page}" (or "{number}") INSIDE it to show the
        /// page number: footer:'Page {page} of {number}'. Text without a placeholder shows NO
        /// page number. "" removes the footer. Null to skip.</param>
        /// <param name="pageNumbers">true = bare page number, centered, in the footer. false =
        /// remove it. Use it when you need no surrounding text; use "{page}" in footer for
        /// "Page X of Y". null = leave untouched.</param>
        /// <returns>Descriptive result message.</returns>
        public string SetHeaderFooter(string? header = null, string? footer = null, bool? pageNumbers = null)
        {
            if (_document == null) return "Error: No document open.";
            try
            {
                var mainPart = _document.MainDocumentPart!;
                var body = GetBody() ?? mainPart.Document.Body;
                if (body == null) return "Error: Document body is empty.";

                var results = new List<string>();
                var sections = body.Elements<SectionProperties>().ToList();

                if (header != null)
                    results.Add(header.Length == 0 ? RemoveSectionPart(mainPart, sections, true) : SetSectionText(mainPart, sections, true, header));
                if (footer != null)
                    results.Add(footer.Length == 0 ? RemoveSectionPart(mainPart, sections, false) : SetSectionText(mainPart, sections, false, footer));

                if (pageNumbers == true)
                {
                    var section = EnsureSections(body)[0];
                    var footerPart = mainPart.FooterParts?.FirstOrDefault() ?? mainPart.AddNewPart<FooterPart>();
                    footerPart.Footer ??= new Footer();
                    if (!footerPart.Footer.Descendants<FieldCode>().Any(f => f.Text?.Contains("PAGE") == true))
                    {
                        // A page-number field needs fldChar begin + code + separate + result + end.
                        footerPart.Footer.Append(new Paragraph(
                            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                            new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }),
                            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                            new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve }),
                            new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
                    }
                    footerPart.Footer.Save();
                    var footerRef = section.GetFirstChild<FooterReference>() ?? section.AppendChild(new FooterReference());
                    footerRef.Type = HeaderFooterValues.Default;
                    footerRef.Id = mainPart.GetIdOfPart(footerPart);
                    results.Add("page numbers");
                    Log.LogStep("DocumentTool.SetHeaderFooter: page numbers added");
                }
                else if (pageNumbers == false)
                {
                    var footerPart = mainPart.FooterParts?.FirstOrDefault();
                    int removed = 0;
                    if (footerPart?.Footer != null)
                        foreach (var para in footerPart.Footer.Elements<Paragraph>().ToList())
                        {
                            var runs = para.Elements<Run>().ToList();
                            for (int i = 0; i < runs.Count; i++)
                            {
                                var chars = runs[i].Descendants<FieldChar>().ToList();
                                if (chars.Count == 0 && runs[i].Descendants<FieldCode>().Any(f => f.Text?.Contains("PAGE") == true))   // legacy orphan code
                                {
                                    runs[i].Remove();
                                    removed++;
                                    continue;
                                }
                                if (chars.Any(f => f.FieldCharType == FieldCharValues.Begin))   // whole begin→end field
                                {
                                    int j = i;
                                    while (j < runs.Count)
                                    {
                                        bool hasEnd = runs[j].Descendants<FieldChar>().Any(f => f.FieldCharType == FieldCharValues.End);
                                        runs[j].Remove();
                                        removed++;
                                        if (hasEnd) break;
                                        j++;
                                    }
                                    break;
                                }
                            }
                        }
                    if (removed > 0) footerPart!.Footer!.Save();
                    results.Add(removed > 0 ? "page numbers removed" : "no page numbers present");
                    Log.LogStep("DocumentTool.SetHeaderFooter: page numbers removed");
                }

                Persist();
                return results.Count > 0 ? $"Header/footer updated: {string.Join(", ", results)}." : "Nothing to set. Pass header, footer, and/or pageNumbers.";
            }
            catch (Exception ex)
            {
                return $"Error: SetHeaderFooter failed. {ex.Message}";
            }
        }

        /// <summary>Inserts a RASTER image file (JPEG/PNG/GIF/BMP/TIFF) — a picture. For data charts
        /// use AddChart instead. Appends at the end, or inserts before the paragraph at index.
        /// Remove it with ParagraphOp(index, "delete") — GetParagraphs() reports type="image".</summary>
        /// <param name="imagePath">Image file path, Unix style relative to the workspace root (e.g. "/folder/image.png").</param>
        /// <param name="index">0-based paragraph index to insert before (from GetParagraphs()). Omit to append at the end.</param>
        /// <param name="width">Width in inches (default 5; height auto 3:4).</param>
        /// <returns>Descriptive result message.</returns>
        public string AddImage(string imagePath, int? index = null, double? width = null) => Run("AddImage", body =>
        {
            // Validate the insertion index FIRST — a rejected index must not leave an
            // orphaned image part (AddImagePart + FeedData) behind in the package.
            var err = InsertIndex(body, index, out var idx);
            if (err != null) return err;

            var resolved = SandboxPath.Resolve(imagePath);
            if (!File.Exists(resolved))
                return $"Error: Image '{imagePath}' not found.";
            // Detect the real format from magic bytes — embedding a PNG as a JPEG part corrupts the docx.
            var format = DetectImageFormat(resolved);
            if (format == null)
                return $"Error: Unsupported or invalid image format for '{imagePath}'. Use JPEG, PNG, GIF, BMP, or TIFF.";

            var mainPart = _document!.MainDocumentPart!;
            var imagePart = mainPart.AddImagePart(format);
            using (var stream = File.OpenRead(resolved))
                imagePart.FeedData(stream);
            var relId = mainPart.GetIdOfPart(imagePart);

            var w = (long)((width ?? 5.0) * 914400);
            var drawing = new Drawing($@"<w:drawing xmlns:w=""{WmlNs}"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
              <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"">
                <wp:extent cx=""{w}"" cy=""{(long)(w * 0.75)}""/>
                <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
                <wp:docPr id=""{int.Parse(NewId())}"" name=""Image""/>
                <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=""1""/></wp:cNvGraphicFramePr>
                <a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id=""0"" name=""Image""/><pic:cNvPicPr/></pic:nvPicPr>
                    <pic:blipFill><a:blip r:embed=""{relId}""/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{w}"" cy=""{(long)(w * 0.75)}""/></a:xfrm><a:prstGeom prst=""rect""/></pic:spPr>
                  </pic:pic>
                </a:graphicData></a:graphic>
              </wp:inline>
            </w:drawing>");

            return Done($"DocumentTool.AddImage: '{resolved}' at index {index?.ToString() ?? "end"}", new Paragraph(new Run(drawing)), idx, $"Image inserted at paragraph position {index?.ToString() ?? "end"}.");
        });

        /// <summary>Adds a NATIVE editable chart (bar/line/pie) in its own paragraph. Flexible input
        /// formats are accepted so the exact JSON shape does not matter.</summary>
        /// <param name="chartType">"bar", "line", or "pie".</param>
        /// <param name="categories">Category labels. JSON array (["Q1","Q2","Q3"]) or comma-separated ("Q1,Q2,Q3").</param>
        /// <param name="series">Series data, ANY of:
        ///   - JSON array of arrays: [[120,180,240],[80,110,150]]
        ///   - JSON array of objects: [{"name":"Revenue","values":[120,180,240]}]
        ///   - comma-separated single series: "120,180,240"
        ///   - semicolon-separated named series: "Revenue:120,180,240;Costs:80,110,150"</param>
        /// <param name="seriesNames">Optional series names (JSON array or comma-separated), used when
        /// <paramref name="series"/> has no names. Defaults to "Series N".</param>
        /// <param name="title">Optional chart title.</param>
        /// <param name="index">Optional 0-based paragraph position (from GetParagraphs()). Omit to append at the end.</param>
        /// <param name="width">Chart width in inches (default 5).</param>
        /// <returns>Descriptive result message.</returns>
        public string AddChart(string chartType, string? categories = null, string? series = null,
            string? seriesNames = null, string? title = null, int? index = null, double? width = null) => Run("AddChart", body =>
        {
            var err = InsertIndex(body, index, out var idx);
            if (err != null) return err;

            var cats = ParseStringArray(categories);
            if (cats == null || cats.Length == 0)
                return "Error: categories cannot be empty. Use [\"Q1\",\"Q2\"] or \"Q1,Q2\".";
            var (values, names) = ParseSeries(series);
            if (values == null || values.Length == 0)
                return "Error: series cannot be empty. Use [[1,2],[3,4]] or [{\"name\":\"Revenue\",\"values\":[1,2]}].";

            var type = chartType.Trim().ToLowerInvariant();
            if (type != "bar" && type != "line" && type != "pie")
                return $"Error: Unknown chart type '{chartType}'. Use 'bar', 'line', or 'pie'.";

            var explicitNames = ParseStringArray(seriesNames);
            for (int i = 0; i < values.Length; i++)
                if (string.IsNullOrEmpty(names[i]))
                    names[i] = explicitNames != null && i < explicitNames.Length
                        ? explicitNames[i]
                        : "Series " + (i + 1);

            var mainPart = _document!.MainDocumentPart!;
            var chartPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
            chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(BuildChartXml(type, title, cats, names, values));
            chartPart.ChartSpace.Save();
            var relId = mainPart.GetIdOfPart(chartPart);

            var w = (long)((width ?? 5.0) * 914400);
            var drawing = new Drawing($@"<w:drawing xmlns:w=""{WmlNs}"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
              <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"">
                <wp:extent cx=""{w}"" cy=""{(long)(w * 0.75)}""/>
                <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
                <wp:docPr id=""{int.Parse(NewId())}"" name=""Chart""/>
                <wp:cNvGraphicFramePr/>
                <a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:chart r:id=""{relId}""/></a:graphicData></a:graphic>
              </wp:inline>
            </w:drawing>");

            return Done($"DocumentTool.AddChart: {type} chart, {values.Length} series, {cats.Length} categories at index {index?.ToString() ?? "end"}", new Paragraph(new Run(drawing)), idx, $"{type} chart added ({values.Length} series × {cats.Length} categories).");
        });

        /// <summary>Parses a list (categories or series names): JSON array, or comma-separated string.
        /// Returns null when empty/unparseable.</summary>
        private static string[]? ParseStringArray(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Trim();
            if (s.StartsWith('['))
            {
                try
                {
                    using var doc = JsonDocument.Parse(s);
                    if (doc.RootElement.ValueKind == JsonValueKind.Array)
                        return doc.RootElement.EnumerateArray()
                            .Select(e => e.ValueKind == JsonValueKind.String ? e.GetString()! : e.GetRawText())
                            .ToArray();
                }
                catch { return new[] { s }; }
            }
            return s.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }

        /// <summary>Parses series into (values, names). Accepts JSON "[[1,2],[3,4]]", JSON
        /// "[{"name":"X","values":[1,2]}]", "1,2,3", or "X:1,2,3;Y:4,5,6". Null values when unparseable.</summary>
        private static (double[][]? Values, string[] Names) ParseSeries(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return (null, Array.Empty<string>());
            s = s.Trim();

            if (s.StartsWith('['))
            {
                try
                {
                    using var doc = JsonDocument.Parse(s);
                    if (doc.RootElement.ValueKind != JsonValueKind.Array) return (null, Array.Empty<string>());
                    var result = new List<double[]>();
                    var names = new List<string>();
                    foreach (var item in doc.RootElement.EnumerateArray())
                    {
                        if (item.ValueKind == JsonValueKind.Array)
                        {
                            result.Add(item.EnumerateArray().Select(JsonToDouble).ToArray());
                            names.Add("");
                        }
                        else if (item.ValueKind == JsonValueKind.Object)
                        {
                            if (!(item.TryGetProperty("values", out var v) || item.TryGetProperty("data", out v)))
                                return (null, Array.Empty<string>());
                            result.Add(v.EnumerateArray().Select(JsonToDouble).ToArray());
                            names.Add(item.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "");
                        }
                        else return (null, Array.Empty<string>());
                    }
                    return (result.ToArray(), names.ToArray());
                }
                catch { /* fall through to textual */ }
            }

            var values = new List<double[]>();
            var namesList = new List<string>();
            foreach (var chunk in s.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                string numbers = chunk;
                string name = "";
                int colon = chunk.IndexOf(':');
                if (colon > 0)
                {
                    name = chunk[..colon].Trim();
                    numbers = chunk[(colon + 1)..].Trim();
                }
                var parts = numbers.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (parts.Length == 0) continue;
                var vals = new double[parts.Length];
                bool ok = true;
                for (int i = 0; i < parts.Length; i++)
                    if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out vals[i])) { ok = false; break; }
                if (!ok) return (null, Array.Empty<string>());
                values.Add(vals);
                namesList.Add(name);
            }
            return (values.Count > 0 ? values.ToArray() : null, namesList.ToArray());
        }

        /// <summary>Reads a JSON number (or numeric string) as double.</summary>
        private static double JsonToDouble(JsonElement e) =>
            e.ValueKind == JsonValueKind.Number ? e.GetDouble()
            : double.TryParse(e.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0;

        /// <summary>Searches for text and returns paragraph indices where it appears.</summary>
        /// <param name="searchText">Text to find.</param>
        /// <returns>JSON array of {paragraphIndex, text, matchCount} objects.</returns>
        public string FindText(string searchText) => Run("FindText", body =>
        {
            if (string.IsNullOrEmpty(searchText)) return "Error: searchText cannot be empty.";
            var results = body.Elements<Paragraph>()
                .Select((p, i) => new { index = i, text = p.InnerText })
                .Where(x => x.text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Select(x => new
                {
                    paragraphIndex = x.index,
                    preview = Truncate(x.text.Trim(), 80),
                    matchCount = (x.text.Length - x.text.Replace(searchText, "", StringComparison.OrdinalIgnoreCase).Length) / searchText.Length
                }).ToList();
            Log.LogStep($"DocumentTool.FindText: '{Truncate(searchText, 40)}' → {results.Count} paragraph(s)");
            return JsonSerializer.Serialize(results);
        });

        /// <summary>Inserts a page break. Appends at the end by default, or inserts before a specific
        /// paragraph when <paramref name="index"/> is provided.</summary>
        /// <param name="index">Optional 0-based paragraph position where to insert the page break (from GetParagraphs()). Omit to append at the end.</param>
        /// <returns>Descriptive result message.</returns>
        public string AddPageBreak(int? index = null) => Run("AddPageBreak", body =>
        {
            var err = InsertIndex(body, index, out var idx);
            return err ?? Done($"DocumentTool.AddPageBreak: at index {index?.ToString() ?? "end"}", new Paragraph(new Run(new Break { Type = BreakValues.Page })), idx, $"Page break inserted at paragraph position {index?.ToString() ?? "end"}.");
        });

        /// <summary>Edits an existing table: add/delete rows, add/delete columns, delete the whole table, or set a cell's text.</summary>
        /// <param name="tableIndex">0-based table index (from GetDocumentInfo()).</param>
        /// <param name="action">"add_row", "delete_row", "add_column", "delete_column", "delete_table", or "set_cell".</param>
        /// <param name="row">Cell values for "add_row" (new row) or "add_column" (new column, one per existing row).</param>
        /// <param name="rowIndex">0-based row index for "set_cell" / "delete_row".</param>
        /// <param name="colIndex">0-based column index for "set_cell" / "delete_column", or insertion position for "add_column" (omitted = end).</param>
        /// <param name="text">New cell text for "set_cell".</param>
        /// <returns>Descriptive result message.</returns>
        public string TableEdit(int tableIndex, string action, string[]? row = null, int? rowIndex = null, int? colIndex = null, string? text = null) => Run("TableEdit", body =>
        {
            var table = GetTableAt(body, tableIndex, out var err);
            if (err != null) return err;

            // Tolerate camelCase ("addRow") and whitespace — LLMs commonly use both forms.
            action = Regex.Replace(action.Trim(), "([a-z0-9])([A-Z])", "$1_$2").ToLowerInvariant();
            switch (action)
            {
                case "add_row":
                    if (row == null || row.Length == 0) return "Error: 'row' (cell values array) is required for 'add_row'.";
                    var tr = new TableRow();
                    foreach (var val in row) tr.Append(new TableCell(new Paragraph(new Run(new Text(val ?? "")))));
                    table!.Append(tr);
                    return Done($"DocumentTool.TableEdit: row added to table {tableIndex}, {row.Length} cells", $"Row added to table {tableIndex} ({row.Length} cells).");

                case "delete_row":
                    if (rowIndex == null) return "Error: 'rowIndex' is required for 'delete_row'.";
                    if (table!.Elements<TableRow>().Count() <= 1)
                        return "Error: Cannot delete the last row. A table must keep at least one row (delete_table to remove it entirely).";
                    var delRow = GetRowAt(table!, tableIndex, rowIndex.Value, out err);
                    if (err != null) return err;
                    delRow!.Remove();
                    return Done($"DocumentTool.TableEdit: row {rowIndex} deleted from table {tableIndex}", $"Row {rowIndex} deleted from table {tableIndex}.");

                case "add_column":
                    var rowsA = table!.Elements<TableRow>().ToList();
                    int colCountA = rowsA.Count > 0 ? rowsA[0].Elements<TableCell>().Count() : 0;
                    int colPos = colIndex ?? colCountA;
                    if (colIndex != null && (colIndex < 0 || colIndex > colCountA))
                        return $"Error: Column {colIndex} out of range. Table {tableIndex} has {colCountA} columns (0-{colCountA}).";
                    UpdateTableGrid(table, +1, colPos);
                    int a = 0;
                    foreach (var r in rowsA)
                    {
                        var value = row != null && a < row.Length ? row[a] ?? "" : "";
                        var newCell = new TableCell(new Paragraph(new Run(new Text(value))));
                        var cells = r.Elements<TableCell>().ToList();
                        if (colPos >= cells.Count) r.Append(newCell); else r.InsertBefore(newCell, cells[colPos]);
                        a++;
                    }
                    return Done($"DocumentTool.TableEdit: column added to table {tableIndex} at {colPos} ({rowsA.Count} cells)", $"Column added to table {tableIndex} at position {colPos} ({rowsA.Count} cells).");

                case "delete_column":
                    if (colIndex == null) return "Error: 'colIndex' is required for 'delete_column'.";
                    var rowsD = table!.Elements<TableRow>().ToList();
                    int colCountD = rowsD.Count > 0 ? rowsD[0].Elements<TableCell>().Count() : 0;
                    if (colIndex < 0 || colIndex >= colCountD)
                        return $"Error: Column {colIndex} out of range. Table {tableIndex} has {colCountD} columns (0-{colCountD - 1}).";
                    foreach (var r in rowsD)
                    {
                        var cells = r.Elements<TableCell>().ToList();
                        if (colIndex.Value < cells.Count) cells[colIndex.Value].Remove();
                    }
                    UpdateTableGrid(table, -1, colIndex.Value);
                    return Done($"DocumentTool.TableEdit: column {colIndex} deleted from table {tableIndex}", $"Column {colIndex} deleted from table {tableIndex}.");

                case "delete_table":
                    table!.Remove();
                    return Done($"DocumentTool.TableEdit: table {tableIndex} deleted", $"Table {tableIndex} deleted.");

                case "set_cell":
                    if (rowIndex == null || colIndex == null || text == null)
                        return "Error: 'rowIndex', 'colIndex' and 'text' are required for 'set_cell'.";
                    var cellRow = GetRowAt(table!, tableIndex, rowIndex.Value, out err);
                    if (err != null) return err;
                    var cell = GetCellAt(cellRow!, rowIndex.Value, colIndex.Value, out err);
                    if (err != null) return err;
                    cell!.RemoveAllChildren<Paragraph>();
                    cell.Append(new Paragraph(new Run(new Text(text))));
                    return Done($"DocumentTool.TableEdit: cell [{rowIndex},{colIndex}] in table {tableIndex} = '{Truncate(text, 40)}'", $"Cell [{rowIndex},{colIndex}] in table {tableIndex} set to '{text}'.");

                default:
                    return $"Error: Unknown action '{action}'. Use 'add_row', 'delete_row', 'add_column', 'delete_column', 'delete_table', or 'set_cell'.";
            }
        });

        /// <summary>Converts a DOCX to PDF (same name, .pdf extension) in the sandbox.
        /// Converts the currently open document when <paramref name="filePath"/> is omitted.
        /// Requires an external PDF engine (Word/LibreOffice); returns an Error otherwise —
        /// it does NOT produce a placeholder file with DOCX bytes.</summary>
        /// <param name="filePath">DOCX path to convert, Unix style relative to the workspace root (e.g. "/folder/file.docx"). Omit to convert the open document.</param>
        /// <returns>Descriptive result message.</returns>
        public string ConvertToPdf(string? filePath = null)
        {
            try
            {
                if (filePath != null)
                {
                    var source = SandboxPath.Resolve(filePath);
                    if (!File.Exists(source)) return $"Error: File '{filePath}' not found.";
                }
                else if (_document == null) return "Error: No document open.";
                else if (string.IsNullOrEmpty(_filePath)) return "Error: No file path. Open a document first.";

                return "Error: ConvertToPdf requires a real PDF engine (Microsoft Word or LibreOffice). " +
                       "The previous placeholder copied DOCX bytes into a .pdf file, which is not a valid PDF.";
            }
            catch (Exception ex)
            {
                return $"Error: ConvertToPdf failed. {ex.Message}";
            }
        }

        /// <summary>Adds a comment to a specific paragraph.</summary>
        /// <param name="paragraphIndex">0-based paragraph index (from GetParagraphs()).</param>
        /// <param name="text">Comment text.</param>
        /// <param name="author">Comment author (defaults to "AI Agent").</param>
        /// <returns>Descriptive result message.</returns>
        public string AddComment(int paragraphIndex, string text, string? author = null) => Run("AddComment", body =>
        {
            var para = GetParagraphAt(body, paragraphIndex, out var err);
            if (err != null) return err;

            var mainPart = _document!.MainDocumentPart!;
            var wpId = NewId();
            var commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments ??= new Comments();
            // The comment body holds ONLY the comment text; the anchor in the main document
            // carries the CommentReference (Word's own structure — a commentReference INSIDE
            // the comment trips the schema validator's semantic reference check).
            commentsPart.Comments.Append(new Comment(
                new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))
            { Id = wpId, Author = author ?? "AI Agent", Date = DateTime.Now });
            commentsPart.Comments.Save();

            para!.AppendChild(new Run()).AppendChild(new CommentReference { Id = wpId });
            return Done($"DocumentTool.AddComment: paragraph {paragraphIndex}, '{Truncate(text, 60)}'", $"Comment added to paragraph {paragraphIndex}.");
        });

        /// <summary>Generates a Table of Contents based on heading styles (Heading1-3).
        /// Inserts at the specified paragraph index, or at the beginning if not specified.
        /// The TOC updates automatically when the document is opened in Word.</summary>
        /// <param name="index">Optional 0-based paragraph position where to insert the TOC (from GetParagraphs()). Omit to insert at the beginning.</param>
        /// <param name="title">Optional title for the TOC section (e.g. "Table of Contents").</param>
        /// <param name="maxLevel">Maximum heading level to include (1-9, default 3).</param>
        /// <returns>Descriptive result message.</returns>
        public string AddTableOfContents(int? index = null, string? title = null, int maxLevel = 3) => Run("AddTableOfContents", body =>
        {
            maxLevel = Math.Clamp(maxLevel, 1, 9);
            int insertIdx = index ?? -1;   // null = beginning

            if (!string.IsNullOrEmpty(title))
            {
                int tpos = insertIdx < 0 ? 0 : Math.Min(insertIdx, body.Elements<Paragraph>().Count());
                InsertAt(body, new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                    new Run(new Text(title) { Space = SpaceProcessingModeValues.Preserve })), tpos);
                insertIdx = Math.Min(tpos + 1, body.Elements<Paragraph>().Count());
            }

            var tocPara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }, new SpacingBetweenLines { Before = "120", After = "120" }),
                new Run(
                    new FieldChar { FieldCharType = FieldCharValues.Begin },
                    new FieldCode($@" TOC \o ""1-{maxLevel}"" \h \z \u ") { Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new RunProperties(new Color { Val = "808080" }), new Text("(Update table of contents in Word)")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            InsertAt(body, tocPara, Math.Max(insertIdx, 0));

            return Done($"DocumentTool.AddTableOfContents: maxLevel={maxLevel}, title='{title ?? "(none)"}'", $"Table of contents added (headings 1-{maxLevel}). Update in Word to populate.");
        });

        /// <summary>Returns all available paragraph styles as a JSON array. Each entry has id and name.</summary>
        /// <returns>JSON string with available styles.</returns>
        public string GetAvailableStyles()
        {
            if (_document == null) return "Error: No document open.";
            try
            {
                var styles = _document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
                if (styles == null) return "[]";
                var result = styles.Descendants<Style>()
                    .Where(s => s.Type == StyleValues.Paragraph)
                    .Select(s => new { id = s.StyleId?.Value ?? "", name = s.StyleName?.Val?.Value ?? s.StyleId?.Value ?? "" })
                    .ToList();
                Log.LogStep($"DocumentTool.GetAvailableStyles: {result.Count} styles");
                return JsonSerializer.Serialize(result);
            }
            catch (Exception ex)
            {
                return $"Error: GetAvailableStyles failed. {ex.Message}";
            }
        }

        /// <summary>Creates a DOCX document from markdown text (preferred for creation) or from a markdown file,
        /// using DocSharp.Markdown's MarkdownConverter.
        /// Supported markdown: headings H1-H6, paragraphs, bold, italic, inline code, fenced code blocks,
        /// block quotes, thematic breaks, ordered/unordered/nested lists, tables, links, and images.
        /// Does NOT support mermaid diagrams. Provide exactly one of <paramref name="markdown"/> or <paramref name="markdownFile"/>.</summary>
        /// <param name="filePath">Output .docx file path, Unix style relative to the workspace root (e.g. "/folder/file.docx").</param>
        /// <param name="markdown">Markdown content to convert. Null if converting a file.</param>
        /// <param name="markdownFile">Path to a markdown file, Unix style relative to the workspace root. Null if passing content.</param>
        /// <returns>Descriptive result message.</returns>
        public string CreateFromMarkdown(string filePath, string? markdown = null, string? markdownFile = null)
        {
            CloseDocument(saveFirst: false);   // release any open handle before overwriting the file
            _filePath = string.Empty;
            try
            {
                var resolved = SandboxPath.Resolve(filePath);
                string mdText;
                if (markdown != null) mdText = markdown;
                else if (markdownFile != null)
                {
                    var mdResolved = SandboxPath.Resolve(markdownFile);
                    if (!File.Exists(mdResolved)) return $"Error: Markdown file '{markdownFile}' not found.";
                    mdText = File.ReadAllText(mdResolved);
                }
                else return "Error: Provide either 'markdown' (content) or 'markdownFile' (path).";

                _fileCreatedThisSession = false;

                var converter = new MarkdownConverter();
                using (var fs = File.Create(resolved))
                {
                    using var doc = converter.ToWordprocessingDocument(MarkdownSource.FromMarkdownString(mdText), fs, WordprocessingDocumentType.Document, append: false);
                    // DocSharp.Markdown uses MDHeading* style ids; remap to standard Heading* so the
                    // DOCX→Markdown converter (AllToMarkdown) recognizes headings.
                    foreach (var p in doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>())
                    {
                        var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        if (styleId != null && styleId.StartsWith("MDHeading", StringComparison.Ordinal))
                            p.ParagraphProperties!.ParagraphStyleId!.Val = "Heading" + styleId["MDHeading".Length..];
                    }
                    doc.Save();
                }

                _document = OpenEditable(resolved);
                _filePath = resolved;
                var versionId = GitSupport.Snapshot(resolved, "DocumentTool create from markdown");

                Log.LogStep($"DocumentTool.CreateFromMarkdown: created '{resolved}' ({mdText.Length} chars md) version='{versionId}'");
                var agentPath = SandboxPath.ToAgent(resolved);
                return versionId != null
                    ? $"Created '{agentPath}' from markdown. New version: {versionId}. (Rollback via GitTool.restore.)"
                    : $"Created '{agentPath}' from markdown.";
            }
            catch (Exception ex)
            {
                return $"Error: CreateFromMarkdown failed. {ex.Message}";
            }
        }

        /// <summary>Converts current DOCX to Markdown via AllToMarkdown's Converter (DocSharp underneath).
        /// Headings, emphasis, lists, tables, code, quotes, links, images are supported.
        /// Table of Contents fields are not rendered: Word populates them on update, and their
        /// content is redundant with the headings already present in the markdown — absence is normal.
        /// Changes are already auto-saved, so the conversion reflects the current state.</summary>
        /// <returns>Markdown representation of the document.</returns>
        public string ToMarkdown()
        {
            if (_document == null) return "Error: No document open.";
            if (string.IsNullOrEmpty(_filePath)) return "Error: No file path. Open a document first.";
            try
            {
                using var stream = new MemoryStream();
                _document.Clone(stream);
                stream.Position = 0;
                var markdown = AllToMarkdown.Converter.ConvertDataToMarkdown(stream, AllToMarkdown.Converter.SupportedFileFormat.docx).TrimEnd();
                Log.LogStep($"DocumentTool.ToMarkdown: converted '{_filePath}' → {markdown.Length} chars md");
                return markdown;
            }
            catch (Exception ex)
            {
                return $"Error: ToMarkdown failed. {ex.Message}";
            }
        }

        // ─── helpers ───────────────────────────────────────────────────────────

        /// <summary>Runs a body operation with the standard guards: no document → Error; exception → "Error: {op} failed."</summary>
        private string Run(string op, Func<Body, string> fn)
        {
            var body = GetBody();
            if (body == null) return "Error: No document open.";
            try { return fn(body); }
            catch (Exception ex) { return $"Error: {op} failed. {ex.Message}"; }
        }

        /// <summary>Logs, persists, and returns the result — the standard tail of every mutator.</summary>
        private string Done(string log, string result)
        {
            Log.LogStep(log);
            Persist();
            return result;
        }

        /// <summary>Inserts el at idx, then logs, persists, and returns the result.</summary>
        private string Done(string log, OpenXmlElement el, int idx, string result)
        {
            InsertAt(GetBody()!, el, idx);
            return Done(log, result);
        }

        /// <summary>Validates an insertion index against the body's paragraph count; returns an error string or null.</summary>
        private static string? InsertIndex(Body body, int? index, out int idx)
        {
            int count = body.Elements<Paragraph>().Count();
            if (index != null && (index < 0 || index > count))
            {
                idx = count;
                return $"Error: Index {index} out of range. Document has {count} paragraphs (0-{count}).";
            }
            idx = Math.Min(index ?? count, count);
            return null;
        }

        /// <summary>Appends el or inserts it before the paragraph at idx. Content appended at the
        /// end must go BEFORE a trailing SectionProperties (sectPr must be the LAST child of the
        /// body — a paragraph after it is schema-invalid and Word renders it in a phantom section).</summary>
        private static void InsertAt(Body body, OpenXmlElement el, int idx)
        {
            var paras = body.Elements<Paragraph>().ToList();
            if (idx >= paras.Count) AppendBody(body, el);
            else body.InsertBefore(el, paras[idx]);
        }

        /// <summary>Appends el as the last content element, keeping any trailing SectionProperties last.</summary>
        private static void AppendBody(Body body, OpenXmlElement el)
        {
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr != null) body.InsertBefore(el, sectPr);
            else body.Append(el);
        }

        /// <summary>Applies font/color run properties (used by SetDocumentFont and FormatParagraph).</summary>
        private static void ApplyRunProps(Run run, string? fontName, int? fontSize, bool? bold, bool? italic, bool? underline, string? colorHex)
        {
            var rp = run.RunProperties ??= new RunProperties();
            if (fontName != null)
            {
                rp.RunFonts ??= new RunFonts();
                rp.RunFonts.Ascii = fontName;
                rp.RunFonts.HighAnsi = fontName;
            }
            if (fontSize != null)
            {
                rp.FontSize ??= new FontSize();
                rp.FontSize.Val = fontSize.Value.ToString();
            }
            if (bold.HasValue) { if (bold.Value) rp.Bold ??= new Bold(); else rp.RemoveChild(rp.Bold); }
            if (italic.HasValue) { if (italic.Value) rp.Italic ??= new Italic(); else rp.RemoveChild(rp.Italic); }
            if (underline.HasValue) { if (underline.Value) rp.Underline ??= new Underline(); else rp.RemoveChild(rp.Underline); }
            if (colorHex != null)
            {
                rp.Color ??= new Color();
                rp.Color.Val = colorHex.TrimStart('#');
            }
        }

        /// <summary>Styled run for hyperlink/cross-reference display text.
        /// rPr child order matters per schema: w:color comes BEFORE w:u.</summary>
        private static Run LinkRun(string text) => new(
            new RunProperties(new Color { Val = "0563C1" }, new Underline { Val = UnderlineValues.Single }),
            new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        /// <summary>Builds the ChartSpace XML for a native chart with embedded data cache.
        /// Chart types: bar (vertical columns), line, pie. c:f references are standard sheet-style
        /// formulas; renderers use the cached values.</summary>
        private static string BuildChartXml(string chartType, string? title, string[] categories, string[] seriesNames, double[][] series)
        {
            var sb = new StringBuilder();
            sb.Append(@"<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><c:chart>");
            if (!string.IsNullOrEmpty(title))
                sb.Append($@"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=""1400"" b=""1""/></a:pPr><a:r><a:rPr lang=""en-US""/><a:t>{XmlEsc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=""0""/></c:title>");

            sb.Append(@"<c:plotArea><c:layout/>");
            var sers = new StringBuilder();
            for (int s = 0; s < series.Length; s++)
                sers.Append(ChartSeriesXml(chartType, s, seriesNames[s], categories, series[s]));

            switch (chartType)
            {
                case "pie":
                    sb.Append(@"<c:pieChart><c:varyColors val=""1""/>").Append(sers).Append("</c:pieChart>");
                    break;
                case "line":
                    sb.Append(@"<c:lineChart><c:grouping val=""standard""/><c:varyColors val=""0""/>")
                      .Append(sers)
                      .Append(@"<c:axId val=""1""/><c:axId val=""2""/></c:lineChart>")
                      .Append(ChartAxesXml());
                    break;
                default:   // bar = vertical columns
                    sb.Append(@"<c:barChart><c:barDir val=""col""/><c:grouping val=""clustered""/><c:varyColors val=""0""/>")
                      .Append(sers)
                      .Append(@"<c:gapWidth val=""150""/>")
                      .Append(@"<c:axId val=""1""/><c:axId val=""2""/></c:barChart>")
                      .Append(ChartAxesXml());
                    break;
            }
            sb.Append(@"</c:plotArea><c:plotVisOnly val=""1""/><c:dispBlanksAs val=""gap""/></c:chart></c:chartSpace>");
            return sb.ToString();
        }

        /// <summary>Series palette (Word-style accent colors), cycled per series index.</summary>
        private static readonly string[] ChartColors = { "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47" };

        /// <summary>One chart series: idx/order + name + spPr fill + categories (strCache) + values (numCache).
        /// A per-series c:spPr fill is REQUIRED for LibreOffice to draw the shapes (bar/line/slice) —
        /// without it the data parses (axes scale from it) but nothing renders.</summary>
        private static string ChartSeriesXml(string chartType, int s, string name, string[] categories, double[] values)
        {
            var color = ChartColors[s % ChartColors.Length];
            var sb = new StringBuilder();
            sb.Append($@"<c:ser><c:idx val=""{s}""/><c:order val=""{s}""/>")
              .Append($@"<c:tx><c:strRef><c:f>Sheet1!$A${s + 1}</c:f><c:strCache><c:ptCount val=""1""/><c:pt idx=""0""><c:v>{XmlEsc(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>")
              .Append($@"<c:spPr><a:solidFill><a:srgbClr val=""{color}""/></a:solidFill><a:ln><a:solidFill><a:srgbClr val=""{color}""/></a:solidFill></a:ln></c:spPr>");
            if (chartType == "line")
                sb.Append($@"<c:marker><c:symbol val=""circle""/><c:size val=""5""/><c:spPr><a:solidFill><a:srgbClr val=""{color}""/></a:solidFill></c:spPr></c:marker>");
            sb.Append($@"<c:cat><c:strRef><c:f>Sheet1!$B$1:$B${categories.Length}</c:f><c:strCache><c:ptCount val=""{categories.Length}""/>");
            for (int i = 0; i < categories.Length; i++)
                sb.Append($@"<c:pt idx=""{i}""><c:v>{XmlEsc(categories[i])}</c:v></c:pt>");
            sb.Append("</c:strCache></c:strRef></c:cat>")
              .Append($@"<c:val><c:numRef><c:f>Sheet1!$C${s + 1}:$C${values.Length}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=""{values.Length}""/>");
            for (int i = 0; i < values.Length; i++)
                sb.Append($@"<c:pt idx=""{i}""><c:v>{values[i].ToString(CultureInfo.InvariantCulture)}</c:v></c:pt>");
            sb.Append("</c:numCache></c:numRef></c:val></c:ser>");
            return sb.ToString();
        }

        /// <summary>Category + value axes wiring (bar/line charts only).</summary>
        private static string ChartAxesXml() =>
            @"<c:catAx><c:axId val=""1""/><c:scaling><c:orientation val=""minMax""/></c:scaling><c:delete val=""0""/><c:axPos val=""b""/><c:crossAx val=""2""/><c:crosses val=""autoZero""/></c:catAx>" +
            @"<c:valAx><c:axId val=""2""/><c:scaling><c:orientation val=""minMax""/></c:scaling><c:delete val=""0""/><c:axPos val=""l""/><c:majorGridlines/><c:numFmt formatCode=""General"" sourceLinked=""0""/><c:crossAx val=""1""/><c:crosses val=""autoZero""/></c:valAx>";

        /// <summary>Escapes a string for XML element text.</summary>
        private static string XmlEsc(string s) => s
            .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

        /// <summary>Parses a length to inches, accepting plain numbers ("1", "0.75") or
        /// unit-suffixed values ("1in", "2.54cm", "25.4mm", "72pt"). Null when unparseable.</summary>
        private static double? ParseInches(string s)
        {
            var m = Regex.Match(s.Trim(), @"^([0-9]*\.?[0-9]+)\s*(in|inch|inches|cm|mm|pt)?$", RegexOptions.IgnoreCase);
            if (!m.Success || !double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                return null;
            return (m.Groups[2].Value ?? "").ToLowerInvariant() switch
            {
                "cm" => v / 2.54,
                "mm" => v / 25.4,
                "pt" => v / 72.0,
                _ => v,
            };
        }

        /// <summary>Sets a header (isHeader=true) or footer paragraph and wires the section reference.
        /// The text may contain "{page}" or "{number}" placeholders, which become a native
        /// page-number field (e.g. "Page {number} of 12").</summary>
        private static string SetSectionText(MainDocumentPart mainPart, List<SectionProperties> sections, bool isHeader, string text)
        {
            var section = sections.Count > 0 ? sections[0] : EnsureSections(mainPart.Document.Body)[0];
            if (isHeader)
            {
                var part = mainPart.HeaderParts?.FirstOrDefault() ?? mainPart.AddNewPart<HeaderPart>();
                part.Header ??= new Header();
                part.Header.RemoveAllChildren<Paragraph>();
                part.Header.Append(SectionTextParagraph(text));
                part.Header.Save();
                var r = section.GetFirstChild<HeaderReference>() ?? section.AppendChild(new HeaderReference());
                r.Type = HeaderFooterValues.Default;
                r.Id = mainPart.GetIdOfPart(part);
                Log.LogStep($"DocumentTool.SetHeaderFooter: header '{Truncate(text, 60)}'");
                return "header";
            }
            var fpart = mainPart.FooterParts?.FirstOrDefault() ?? mainPart.AddNewPart<FooterPart>();
            fpart.Footer ??= new Footer();
            // Setting footer text must NOT wipe an existing page-number field: the contract is
            // "pass only what you want to change; omitted ones are left untouched", and page
            // numbers live in the footer part.
            foreach (var p in fpart.Footer.Elements<Paragraph>()
                .Where(p => !p.Descendants<FieldCode>().Any(f => f.Text?.Contains("PAGE") == true)).ToList())
                p.Remove();
            fpart.Footer.Append(SectionTextParagraph(text));
            fpart.Footer.Save();
            var fr = section.GetFirstChild<FooterReference>() ?? section.AppendChild(new FooterReference());
            fr.Type = HeaderFooterValues.Default;
            fr.Id = mainPart.GetIdOfPart(fpart);
            Log.LogStep($"DocumentTool.SetHeaderFooter: footer '{Truncate(text, 60)}'");
            return "footer";
        }

        /// <summary>Builds a header/footer paragraph, converting "{page}"/"{number}" placeholders
        /// into a native PAGE field (fldChar begin + code + separate + result + end).</summary>
        private static Paragraph SectionTextParagraph(string text)
        {
            var para = new Paragraph();
            foreach (var part in Regex.Split(text, @"(\{page\}|\{number\})", RegexOptions.IgnoreCase))
            {
                if (part.Equals("{page}", StringComparison.OrdinalIgnoreCase) ||
                    part.Equals("{number}", StringComparison.OrdinalIgnoreCase))
                {
                    para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                                new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }),
                                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                                new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve }),
                                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
                }
                else if (part.Length > 0)
                {
                    para.Append(new Run(new Text(part) { Space = SpaceProcessingModeValues.Preserve }));
                }
            }
            return para;
        }

        /// <summary>Removes header (isHeader=true) or footer references and parts from all sections.</summary>
        private static string RemoveSectionPart(MainDocumentPart mainPart, List<SectionProperties> sections, bool isHeader)
        {
            int removed = 0;
            foreach (var s in sections)
                removed += isHeader ? s.Elements<HeaderReference>().Count() : s.Elements<FooterReference>().Count();
            foreach (var s in sections)
                if (isHeader)
                {
                    foreach (var r in s.Elements<HeaderReference>().ToList()) r.Remove();
                }
                else
                {
                    foreach (var r in s.Elements<FooterReference>().ToList()) r.Remove();
                }
            int parts = 0;
            if (isHeader)
            {
                var hs = mainPart.HeaderParts?.ToList() ?? new List<HeaderPart>();
                parts = hs.Count;
                foreach (var p in hs) mainPart.DeletePart(p);
            }
            else
            {
                var fs = mainPart.FooterParts?.ToList() ?? new List<FooterPart>();
                parts = fs.Count;
                foreach (var p in fs) mainPart.DeletePart(p);
            }
            var label = isHeader ? "header" : "footer";
            Log.LogStep($"DocumentTool.SetHeaderFooter: {label} removed");
            return removed > 0 || parts > 0 ? $"{label} removed" : $"no {label} present";
        }

        private Body? GetBody() => _document?.MainDocumentPart?.Document?.Body;

        /// <summary>Classifies a paragraph's content: "chart" (native chart drawing), "image" (picture), else "text".</summary>
        private static string ParagraphType(Paragraph p)
        {
            var uris = p.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().Select(g => g.Uri?.Value ?? "").ToList();
            if (uris.Any(u => u.Contains("/chart"))) return "chart";
            if (uris.Any(u => u.Contains("/picture"))) return "image";
            return "text";
        }

        private static Paragraph? GetParagraphAt(Body body, int index, out string? error)
        {
            var paras = body.Elements<Paragraph>().ToList();
            error = index < 0 || index >= paras.Count
                ? $"Error: Index {index} out of range. Document has {paras.Count} paragraphs (0-{paras.Count - 1})."
                : null;
            return error == null ? paras[index] : null;
        }

        private static Table? GetTableAt(Body body, int tableIndex, out string? error)
        {
            var tables = body.Elements<Table>().ToList();
            error = tableIndex < 0 || tableIndex >= tables.Count
                ? (tables.Count == 0 ? "Error: No tables in the document."
                                     : $"Error: Table index {tableIndex} out of range. Document has {tables.Count} tables (0-{tables.Count - 1}).")
                : null;
            return error == null ? tables[tableIndex] : null;
        }

        private static TableRow? GetRowAt(Table table, int tableIndex, int rowIndex, out string? error)
        {
            var rows = table.Elements<TableRow>().ToList();
            error = rowIndex < 0 || rowIndex >= rows.Count
                ? $"Error: Row {rowIndex} out of range. Table {tableIndex} has {rows.Count} rows (0-{rows.Count - 1})."
                : null;
            return error == null ? rows[rowIndex] : null;
        }

        private static TableCell? GetCellAt(TableRow row, int rowIndex, int colIndex, out string? error)
        {
            var cells = row.Elements<TableCell>().ToList();
            error = colIndex < 0 || colIndex >= cells.Count
                ? $"Error: Column {colIndex} out of range. Row {rowIndex} has {cells.Count} cells (0-{cells.Count - 1})."
                : null;
            return error == null ? cells[colIndex] : null;
        }

        private static List<SectionProperties> EnsureSections(Body body)
        {
            var sections = body.Elements<SectionProperties>().ToList();
            if (sections.Count == 0)
            {
                body.Append(new SectionProperties());
                sections = body.Elements<SectionProperties>().ToList();
            }
            return sections;
        }

        /// <summary>Random id for bookmarks, comments, drawings. Shared locked Random: per-call instances
        /// seeded from the same tick produce duplicate sequences.</summary>
        private static string NewId()
        {
            lock (_idLock) return _idRandom.Next(1, 99999).ToString();
        }

        /// <summary>Ensures the numbering part defines bullet (numId 1) and decimal (numId 2), appending them
        /// when the part exists but lacks them (DocSharp.Markdown / external docs define only their own ids).</summary>
        private static void EnsureNumberingPart(MainDocumentPart mainPart)
        {
            var numPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            var numbering = numPart.Numbering ??= new Numbering();
            var existingIds = numbering.Elements<NumberingInstance>().Select(n => n.NumberID?.Value).ToHashSet();
            if (existingIds.Contains(1) && existingIds.Contains(2)) return;

            var existingAbs = numbering.Elements<AbstractNum>().Select(a => a.AbstractNumberId?.Value).ToHashSet();
            int next = 0;
            while (existingAbs.Contains(next)) next++;
            int bulletAbs = next++;
            while (existingAbs.Contains(next)) next++;
            int decimalAbs = next;

            // Schema order: ALL abstractNum elements come before ANY num (NumberingInstance).
            var newAbs = new List<AbstractNum>();
            var newNums = new List<NumberingInstance>();
            if (!existingIds.Contains(1))
            {
                newAbs.Add(new AbstractNum(
                    new Level { LevelIndex = 0, StartNumberingValue = new StartNumberingValue { Val = 1 },
                        NumberingFormat = new NumberingFormat { Val = NumberFormatValues.Bullet }, LevelText = new LevelText { Val = "\u2022" } })
                { AbstractNumberId = bulletAbs });
                newNums.Add(new NumberingInstance(new AbstractNumId { Val = bulletAbs }) { NumberID = 1 });
            }
            if (!existingIds.Contains(2))
            {
                newAbs.Add(new AbstractNum(
                    new Level { LevelIndex = 0, StartNumberingValue = new StartNumberingValue { Val = 1 },
                        NumberingFormat = new NumberingFormat { Val = NumberFormatValues.Decimal }, LevelText = new LevelText { Val = "%1." } })
                { AbstractNumberId = decimalAbs });
                newNums.Add(new NumberingInstance(new AbstractNumId { Val = decimalAbs }) { NumberID = 2 });
            }
            foreach (var a in newAbs) numbering.Append(a);
            foreach (var n in newNums) numbering.Append(n);
            numPart.Numbering.Save();
        }

        /// <summary>Ensures the styles part defines Normal, Title and Heading1-9 (headings are resolved
        /// through the styles part by AllToMarkdown/DocSharp). No-op when the part already exists.</summary>
        private static void EnsureStylesPart(MainDocumentPart mainPart)
        {
            if (mainPart.StyleDefinitionsPart?.Styles != null) return;
            var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Default = true, Type = StyleValues.Paragraph, StyleId = "Normal" },
                new Style(new StyleName { Val = "Title" }) { Type = StyleValues.Paragraph, StyleId = "Title" });
            for (int i = 1; i <= 9; i++)
                styles.Append(new Style(new StyleName { Val = $"Heading {i}" }) { Type = StyleValues.Paragraph, StyleId = $"Heading{i}" });
            stylesPart.Styles = styles;
            stylesPart.Styles.Save();
        }

        /// <summary>Adds/removes a tblGrid column when the table has an explicit grid; otherwise nothing to update.</summary>
        private static void UpdateTableGrid(Table table, int delta, int position)
        {
            var grid = table.Elements<TableGrid>().FirstOrDefault();
            if (grid == null) return;
            var cols = grid.Elements<GridColumn>().ToList();
            if (delta > 0)
            {
                var anchor = position < cols.Count ? cols[position] : null;
                if (anchor != null) grid.InsertBefore(new GridColumn(), anchor); else grid.Append(new GridColumn());
            }
            else if (position < cols.Count) cols[position].Remove();
        }

        /// <summary>Persists pending in-memory changes to disk and versions the new content
        /// (rollback is centralized in GitTool.restore).</summary>
        private void Persist()
        {
            if (_document == null || string.IsNullOrEmpty(_filePath)) return;
            _document.MainDocumentPart?.Document?.Save();
            _document.Save();   // package-level save flushes the stream
            GitSupport.Snapshot(_filePath, "DocumentTool save");
        }

        /// <summary>Opens the document for editing with FileShare.Read so concurrent readers
        /// (ToMarkdown) can access the file while open. The stream is owned separately
        /// because WordprocessingDocument.Open(stream) does NOT dispose it.</summary>
        private WordprocessingDocument OpenEditable(string path)
        {
            CloseDocument(saveFirst: false);
            _fileStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            return WordprocessingDocument.Open(_fileStream, true);
        }

        /// <summary>Detects the image content type from magic bytes; null for unknown/non-image files.</summary>
        private static string? DetectImageFormat(string path)
        {
            try
            {
                using var fs = File.OpenRead(path);
                if (fs.Length < 4) return null;
                Span<byte> h = stackalloc byte[8];
                int n = fs.Read(h);
                if (n >= 3 && h[0] == 0xFF && h[1] == 0xD8 && h[2] == 0xFF) return "image/jpeg";
                if (n >= 8 && h[0] == 0x89 && h[1] == 0x50 && h[2] == 0x4E && h[3] == 0x47) return "image/png";
                if (n >= 4 && h[0] == 0x47 && h[1] == 0x49 && h[2] == 0x46 && h[3] == 0x38) return "image/gif";
                if (n >= 2 && h[0] == 0x42 && h[1] == 0x4D) return "image/bmp";
                if (n >= 4 && ((h[0] == 0x49 && h[1] == 0x49 && h[2] == 0x2A && h[3] == 0x00) ||
                               (h[0] == 0x4D && h[1] == 0x4D && h[2] == 0x00 && h[3] == 0x2A))) return "image/tiff";
                return null;
            }
            catch { return null; }
        }

        /// <summary>
        /// Generates a complete, well-structured PDF document (report, analisi, relazione tecnica, verbale,
        /// studio di caso, ecc.) from a set of context files — instead of building the document manually.
        /// The document pipeline (CreateDocument) produces a full PDF with Mermaid diagrams plus the Markdown
        /// and a ".info" manifest listing the cited documents. The context files are the sole source when
        /// judged sufficient; otherwise internet-retrievable topics are enriched with web sources, while
        /// internal-only topics with insufficient context fail with InsufficientSources.
        /// The call is LONG-RUNNING: when the result starts with "ProcessingInProgress" the generation
        /// continues in the background and the outcome arrives later as an unsolicited TaskCompleted event —
        /// inform the user of the in-progress state and wait for that event before reporting a final result.
        /// </summary>
        /// <param name="contextFiles">Paths (Unix style, e.g. "/folder/file.docx" or "/AIChatAttachments/...")
        /// of the files to use as source material, discovered with FileTool.file_search or listed in the
        /// attachments notice. Non-Markdown files are resolved to their converted Markdown when available.</param>
        /// <param name="destinationFolder">Output folder (Unix style) where the PDF, Markdown and .info files
        /// are written.</param>
        /// <param name="documentType">Type of document to produce: report, resoconto, analisi, sintesi,
        /// relazione tecnica, verbale, studio di caso, monografia, saggio, guida, manuale, rassegna, parere,
        /// proposta progettuale, memoria, ecc.</param>
        /// <param name="subject">The subject matter the document must cover.</param>
        /// <param name="format">Output format. Only "pdf" is supported.</param>
        /// <returns>Either the generated PDF path (sandbox form) when the run completed synchronously, a
        /// "ProcessingInProgress: ... task id N" notice when the generation runs in the background (the
        /// TaskCompleted event will follow), or an "Error: ..." message (unsupported format, missing files,
        /// context too large). Context-too-large errors report the exact numbers — total estimated tokens,
        /// the model's context window and the per-file token estimates — so the agent can trim the context
        /// and retry.</returns>
        public string CreateDocumentFromContext(List<string> contextFiles, string destinationFolder, string documentType, string subject, string format = "pdf")
        {
            if (!string.Equals(format, "pdf", StringComparison.OrdinalIgnoreCase))
                return "Error: only 'pdf' output is supported. Call with format='pdf'.";
            if (contextFiles == null || contextFiles.Count == 0)
                return "Error: contextFiles is empty. Find relevant files with FileTool.file_search and pass their paths.";
            if (string.IsNullOrWhiteSpace(documentType) || string.IsNullOrWhiteSpace(subject))
                return "Error: documentType and subject are required.";

            string outputHostDir;
            try { outputHostDir = SandboxPath.Resolve(destinationFolder); }
            catch (UnauthorizedAccessException ex) { return $"Error: {ex.Message}"; }

            var mdPaths = new List<string>();
            // Build the source→shadow map ONCE: resolving per file would re-read every /.md file for every
            // input (O(K×N)); with the shared map each lookup is O(1).
            var sourceToMd = BuildSourceToMdMap();
            foreach (var file in contextFiles)
            {
                string hostPath;
                try { hostPath = SandboxPath.Resolve(file); }
                catch (UnauthorizedAccessException ex) { return $"Error: {ex.Message}"; }
                if (!File.Exists(hostPath))
                    return $"Error: file '{file}' not found in the workspace.";
                var md = ResolveContextMarkdown(hostPath, sourceToMd);
                if (md == null)
                    return $"Error: cannot obtain Markdown for '{file}'. Pass .md files or files indexed by file_search.";
                mdPaths.Add(md);
            }

            // The heavy work runs in the background via the standard task registry: this call returns the
            // in-progress notice immediately, and the orchestrator either waits (sync mode) or delivers the
            // completion event on the next run (responsive mode). CreateDocument stays synchronous inside the
            // task — the registry provides the Task.Run layer.
            // Context-size handling happens INSIDE CreateDocument.CreateInternal, not here: oversized files
            // are auto-summarized and the set is reduced to the relevant documents that fit the window
            // (DocumentsSupport.FilterContextByRelevance). The agent is never asked to trim the context —
            // only a residual ContextTooLarge (extreme case) carries exact numbers via the callback.
            var dto = new AllowedPropertiesData { DocumentType = documentType, Subject = subject };
            CreateDocument.ContextTooLargeInfo? tooLarge = null;
            // The per-conversation registry comes from the ambient (set by the orchestrator around tool
            // dispatch); outside an orchestrator (tests) the tool owns a throwaway instance assigned to
            // AsyncTaskRegistry so the caller can wait on it.
            var registry = AgentTaskRegistry.Current ?? (AsyncTaskRegistry = new AgentTaskRegistry());
            var taskId = registry.Start("create_document_from_context",
                () => DescribeCreationResult(
                    (SimulateCreateDocument?.Invoke(dto, mdPaths, outputHostDir)
                     ?? CreateDocument.Create(dto, supportingDocuments: mdPaths, outputDirectory: outputHostDir, onContextTooLarge: info => tooLarge = info)),
                    outputHostDir, tooLarge));
            return $"ProcessingInProgress: 'create_document_from_context' running, task id {taskId}. An update will arrive when it completes.";
        }

        /// <summary>
        /// Test seam (harnesses only, see CreateDocumentAgent.Tests): replaces the heavy
        /// <see cref="CreateDocument.Create"/> call inside <see cref="CreateDocumentFromContext"/> so the
        /// agent orchestration flow (attachment discovery → path resolution → task registry → completion
        /// event) can be verified end-to-end without paying the full LLM document-generation cost.
        /// The substitute MUST produce the same observable artifacts as the real pipeline (PDF + ".info"
        /// manifest) for the outcome-based checks to remain valid. Null (default) = real generation.
        /// </summary>
        public static Func<AllowedPropertiesData, List<string>?, string?, DocumentCreationResult>? SimulateCreateDocument { get; set; }

        /// <summary>Builds the "source path → shadow markdown" map from the /.md directory frontmatter ONCE
        /// per call (FileTool uses the same layout). Hoisted out of the per-file loop: resolving each context
        /// file independently would re-enumerate and re-read every shadow file for every input (O(K×N)).</summary>
        private static Dictionary<string, string>? BuildSourceToMdMap()
        {
            var docsPath = Setup.DocumentsPath;
            if (string.IsNullOrWhiteSpace(docsPath)) return null;
            var mdDir = Path.Combine(docsPath, RagDocumentProcessor.MarkdownDirectoryName);
            if (!Directory.Exists(mdDir)) return null;
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in Directory.EnumerateFiles(mdDir, "*.md"))
            {
                try
                {
                    var lines = File.ReadAllLines(f);
                    if (lines.Length >= 2 && lines[0].Trim() == "---")
                    {
                        var srcLine = lines[1].TrimStart();
                        if (srcLine.StartsWith("source:", StringComparison.OrdinalIgnoreCase))
                            map[UnixPath.Normalize(srcLine.Substring(srcLine.IndexOf(':') + 1).Trim())] = f;
                    }
                }
                catch { }
            }
            return map;
        }

        /// <summary>Resolves a context file to its Markdown: .md files are used as-is; other files are looked
        /// up in the /.md shadow directory by their "source:" frontmatter (the layout produced for indexed
        /// documents and chat attachments). Returns null when no Markdown is available.</summary>
        private static string? ResolveContextMarkdown(string hostPath, Dictionary<string, string>? sourceToMd)
        {
            if (string.Equals(Path.GetExtension(hostPath), ".md", StringComparison.OrdinalIgnoreCase))
                return hostPath;
            var docsPath = Setup.DocumentsPath;
            if (string.IsNullOrWhiteSpace(docsPath) || sourceToMd == null) return null;
            var agentPath = UnixPath.Normalize(SandboxPath.ToAgent(hostPath));
            return sourceToMd.TryGetValue(agentPath, out var mdFullPath) && File.Exists(mdFullPath) ? mdFullPath : null;
        }

        /// <summary>Maps the document pipeline result to the narrative text the agent receives as the
        /// "Result:" of the TaskCompleted event. On success the generated PDF is located in the output
        /// folder directly (name lookup, not reconstruction) so the reported path always matches the file.
        /// On <see cref="DocumentCreationResult.ContextTooLarge"/> the <paramref name="tooLarge"/> diagnostics
        /// (captured via the Create callback) are included: total estimated tokens, the model's context window
        /// and the per-file breakdown — the numbers the agent needs to trim the context and retry.</summary>
        private static string DescribeCreationResult(DocumentCreationResult result, string outputHostDir, CreateDocument.ContextTooLargeInfo? tooLarge = null)
        {
            switch (result)
            {
                case DocumentCreationResult.Success:
                    var pdf = Directory.GetFiles(outputHostDir, "*.pdf").OrderByDescending(f => new FileInfo(f).LastWriteTimeUtc).FirstOrDefault();
                    return $"Success — document generated at {SandboxPath.ToAgent(pdf ?? outputHostDir)}";
                case DocumentCreationResult.InsufficientSources:
                    return "Error: the provided context is not sufficient and the topic appears to be internal/private material that cannot be retrieved online. Enrich the context files (more or better documents) and retry, or give up.";
                case DocumentCreationResult.ContextTooLarge:
                    var perFile = tooLarge != null && tooLarge.Files.Count > 0
                        ? string.Join(", ", tooLarge.Files.Select(f => $"{SandboxPath.ToAgent(f.File)} ≈ {f.Tokens:N0} tokens"))
                        : "(details unavailable)";
                    var total = tooLarge?.EstimatedTokens;
                    var win = tooLarge?.ContextWindow;
                    return $"Error: the combined context is too large even after summarization to fit the model window" +
                           (total is int t && win is int w ? $" (estimated {t:N0} tokens of a {w:N0}-token window)" : "") +
                           $". Per file: {perFile}. Reduce the number or size of the context files (keep only the most significant, total well under {win?.ToString("N0") ?? "the window"}) and retry.";
                case DocumentCreationResult.InvalidSubject:
                    return "Error: the subject could not be recognized/translated. Use a clear, well-formed subject.";
                case DocumentCreationResult.LlmUnavailable:
                    return "Error: the LLM provider is temporarily unavailable. The document was not generated; retry later.";
                default:
                    return $"Error: unexpected failure ({result}). Check the log and retry.";
            }
        }

        private static string Truncate(string value, int maxLen)
            => value.Length <= maxLen ? value : value[..maxLen] + "...";
    }
}
