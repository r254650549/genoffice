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

XLSX runtime extraction (2026-10-13; initial extraction checkpoint):

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

Sheets viewport/context-menu fixes (2026-10-13):

- Grow the existing virtual grid from its visible right/bottom edge, including
  initial layout and window resizing. Wide views at A1 no longer end at Z;
  growth remains a view mutation outside document edits and undo history.
- Distinguish the footer by its sheet tablist. Both the canvas and footer use
  `data-range-selector`, so that attribute alone opened the statistics popup
  together with the cell context menu.
- Existing grid-growth and menu suites: 19 tests pass. Parent browser evidence
  verifies AA-column editing and a single cell menu on a wide viewport.

Sheets Web host checkpoint and clean-close fixes (2026-10-13):

- Explicit `OfficeHost` bindings replace renderer preload calls. The complete
  Sheets renderer, journal, native operations and undo stack remain upstream
  implementations. Electron explicitly supplies the desktop host; existing
  renderer test fixtures now configure the same host interface.
- A local-socket workbook service runs beside the original Rust sidecar in the
  user's existing execution environment. Requests enforce browser/document
  ownership, retain mutation receipts and release closed sessions. Idle reads
  are not cached as mutation results. Open sessions survive ordinary idle time;
  the enclosing execution environment still owns its lifecycle and limits.
- First explicit save uses native `save-as`, including an untouched blank
  workbook. Recovery serialization of undo-to-pristine retains the native
  validation rules. Engine reconnection uses verified immutable input bytes;
  unknown mutation outcomes and lost chunk transfers are not replayed.
- New-but-unnamed does not mean dirty. Close commits an active cell before
  testing the journal; the outer host checks dirty before flushing. DOCX no
  longer serializes a blank recovery copy merely because it has no saved file.
- Browser acceptance in the parent proves blank XLSX explicit cloud save and
  download, zero-edit XLSX/DOCX direct close, and canceling explicit Save As
  followed by direct close. Agent, full lifecycle and compatibility acceptance
  remain pending. The old already-open editor may need reopening to load this
  static Web build; its existing draft is not discarded by these changes.
- Sheets typecheck passes. Browser-independent Sheets suites pass with native
  Windows-sidecar suites and platform-specific atomic-file fault tests excluded;
  the real Linux native-engine verification is recorded separately. The exact
  test counts and exclusions are recorded in the parent checkpoint.
