# DocumentTool — Document (DOCX) agent tool

Document (DOCX) operations for LLM agents: open/create, paragraphs, tables,
headers/footers, charts, images. Derives from `AIOrchestrator.API.BaseAgentTool`.

## What it is

A **plugin** for hosts built on [Graphene.AIOrchestrator](https://github.com/Graphene-Lab/AIOrchestrator)
(e.g. AIOffice): the assembly is loaded dynamically from the host's `Tools/` folder
at startup (or hot-added via a filesystem watcher) and its public methods become
LLM-callable tools through the `UISupportGeneric.Analyzer.GetToolDefinitions`
reflection pipeline.

## Usage

- **Host with dynamic loading** (AIOffice): drop `DocumentTool.dll` + `DocumentTool.xml`
  into the host's `Tools/` directory. The host scans the folder at startup and on
  new files (30 s debounce).
- **Host with static reference**: `ProjectReference` the project (or
  `PackageReference Graphene.DocumentTool`) and register
  `typeof(AIOrchestrator.API.DocumentTool)` in the agent types array passed to
  `AgentHarness.ExecuteAction`.

## Build

```
dotnet build -c Release
```

Compiled as **AnyCPU, RID-neutral**: the same `.dll` runs on Linux, Windows, macOS
and iOS. The XML documentation file (`DocumentTool.xml`) is generated and must ship next
to the assembly — `UISupportGeneric` reads method/parameter descriptions from it.

## NuGet

`Graphene.DocumentTool`, date-versioned `1.yy.MM.dd`, published by
`.github/workflows/publish.yml` when a `v*` tag is pushed. No secret is involved: the workflow mints
a short-lived credential through trusted publishing (`NuGet/login@v1` plus
`permissions: id-token: write`), so a plain `master` push publishes nothing.
Local pack pushes with the `NuGetApiKey` env var; skip with `-p:SkipNuGetPush=true`.

## License

See [LICENSE.md](LICENSE.md).
