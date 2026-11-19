# DocumentTool — notes for coding agents

> Read this file first. This repo is one tool plugin of the AIOffice solution: it is compiled on
> its own and loaded by the hosts from `Tools/DocumentTool/` at runtime.

## What this is

Word (DOCX) documents for the agent: open/create, paragraphs, tables, headers and footers,
charts, images.

**The folder name lies.** The repository folder is `WordTool` (legacy), the project is
`DocumentTool.csproj`, the assembly is `DocumentTool.dll`, and the host folder is
`Tools/DocumentTool`.

## Read before you edit

| If you touch… | Read first |
|---|---|
| anything | [PLUGIN-GUARDRAILS](../AIOrchestrator/docs-dev/PLUGIN-GUARDRAILS.md) — the contract shared by the 12 plugins; the file lives in the sibling AIOrchestrator repo (private), reachable from the solution checkout |
| a tool method or its XML docs | [AGENT_TOOLS_GUIDE](../AIOrchestrator/API/AGENT_TOOLS_GUIDE.md) |
| the document model | `README.md` |

## Non-negotiables

1. The class derives from `BaseAgentTool` and keeps `[ToolPurpose(ToolPurpose.Documents)]`.
2. Public methods and their XML docs **are** the agent's surface: no public method without docs,
   and a public method is never renamed (the model calls it by that name).
3. AnyCPU, no `RuntimeIdentifier`, `GenerateDocumentationFile` stays on.
4. The plugin folder never carries `AIOrchestrator.dll`; no state is written into `Tools/`.
5. A new metadata goes in a class attribute (or a `virtual` member with a default) — never an
   `abstract` member: the dlls already shipped must keep loading.

## Where things are

- `DocumentTool.cs` — the tool class; its public methods are what the model can call
- `CreateDocumentAgent.Tests/`, `DocumentTool.Tests/`, `UnixPath.Tests/` — nested test projects,
  excluded from the library compile
- `DocumentTool.csproj` — the packaging, the version, the assembly metadata

## Shipping

Publishes a GitHub release zip **and** a NuGet package; AgentBridge ships it in every release and
AIOffice builds it into `Tools/DocumentTool` on every dev build.
