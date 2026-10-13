# KompAgent Office host branch

Upstream: https://github.com/genspark-ai/genoffice.git
Origin: https://github.com/r254650549/genoffice.git
Baseline: `fc754f7a90619289942602ebcce68bd57a90515c`.
Branch: `codex/kompagent-office-host`.

This repository retains the complete upstream Git history. KompAgent includes
it as a Git submodule and pins an adapted commit, rather than copying source
files or changing them during a build. Original licenses and notices remain.
The integration must consume the complete upstream editor renderer and its
native document operations and save pipeline, following Happy Friday Lite's
whole-editor host assembly. KompAgent replaces the Electron host boundary for
the Web; it must not substitute a reduced editor or a custom formatting toolbar.
`ee/`, Electron main/preload and a second account/model configuration are outside
the planned Web/runtime bundles. The DOCX Web host now mounts the complete renderer; product acceptance remains tracked in the parent repository.

Adaptations (2026-10-13):

- Explicit `OfficeHost extends DesktopApi` replaces renderer preload references.
  Electron injects its original preload adapter. No runtime monkey-patching or
  build-time source replacement is used.
- `web.html`, `web-main.tsx`, `web-host.ts`, `web-protocol.ts` and
  `vite.web.config.ts` build the complete App, Ribbon, document lifecycle,
  pagination and native DOCX save pipeline for an isolated same-origin iframe.
- Native MCP commands still execute through upstream `executeTool`, including
  block context, undo and OOXML serialization. Revision checks and snapshot
  publication bridge the live document to KompAgent's existing Agent.
- Web AI actions open KompAgent's current conversation; the upstream AiPanel and
  model/provider settings are not mounted by this Web host.
- Zero-sized hidden iframe measurements cannot enter width-fit zoom.
- Draft publication retries transient failures with bounded backoff. File saves
  are not automatically replayed. Read-only views do not publish drafts.
- File read/write, tab, image picker, downloads and recovery use parent-host RPC.
  Desktop-only integrations explicitly report unavailable Web adapters.
- `npm ci --ignore-scripts` installs the pinned package-lock dependencies;
  `node node_modules/vite/bin/vite.js build --config apps/docs/vite.web.config.ts`
  builds DOCX Web assets. The parent build invokes this command.
- DOCX Web build and full Docs TypeScript check passed. Parent evidence records
  real browser input, real-model native commands, save/reopen/download, and
  remaining compatibility and fault-recovery acceptance.

To update upstream, fetch the `upstream` remote, review and merge a selected
upstream commit into this branch, resolve host-adapter conflicts, run the DOCX
round-trip/browser checks, commit here, then update the parent submodule pointer.
Do not change the parent pointer before the adapted child commit is available
from its configured origin.

DOCX tool closure (2026-10-13):

- The current-document bridge exposes the 30 native editing/read tools plus
  discovery, save, undo and redo. Tool schemas and all 23 batch operation
  signatures come directly from the upstream definitions.
- `mcp-document-tools.ts` passes App-owned comments, notes, page setup,
  headers/footers, styles, watermark and protection contexts to `executeTool`.
  Frozen selection tokens prevent cursor movement from retargeting edits.
- Command execution and autosave coordinate through the explicit OfficeHost;
  snapshot publication completes before the next queued command is handled.
- Image insertion also accepts native data URLs. KompAgent can resolve a
  permission-checked immutable cloud image into this format.
- Native undo does not roll back every separate document metadata store. The
  tool catalog states that boundary. The second-agent writer, search, media
  services and cross-format creation are not exposed by this document bridge.
- Native bridge, comments, revisions, charts and style/floating tests: 46 pass.
  Parent evidence includes real-model edits, native undo/redo, cloud save/reopen,
  tracked-change accept/reject and DOCX XML round-trip checks.

XLSX runtime extraction (2026-10-13, Web integration still pending):

- Move `writeWorkbookTo` unchanged out of `sheets-main.ts` into
  `workbook-save.ts`. Electron imports this same function; there is no second
  spreadsheet writer or snapshot-to-XLSX replacement.
- `WorkbookRuntime` owns native sidecar sessions and immutable input snapshots,
  validates range/recalc/save/chunk requests with the existing schemas, and
  uses the original Rust sidecar and OOXML gateway. A prepared save opens a new
  snapshot and keeps the previous session intact until the outer host confirms
  publication. The outer host must enforce user/browser ownership and release
  sessions; this module is not a public network service.
- Rust 1.95.0, Cargo.lock, Node 24 and the source fork are fixed by the parent
  build script. The runtime bundle excludes Electron and enterprise code.
- A real Linux-container integration test covers native blank workbook creation,
  ranges, formulas, formatting, multiple sheets, immutable save, reopen, range
  limits, expired sessions and untouched OOXML parts. It does not constitute
  browser, Agent or user-execution-environment acceptance.
