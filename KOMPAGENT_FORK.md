# KompAgent Office host branch

Upstream: https://github.com/genspark-ai/genoffice.git
Origin: https://github.com/r254650549/genoffice.git
Baseline: `fc754f7a90619289942602ebcce68bd57a90515c`.
Branch: `codex/kompagent-office-host`.

This repository retains the complete upstream Git history. KompAgent includes
it as a Git submodule and pins an adapted commit, rather than copying source
files or changing them during a build. Original licenses and notices remain.
The KompAgent build consumes the DOCX engine and editor model; `ee/`, Electron
main/preload, the upstream Agent, accounts and model settings are not included
in the Web/runtime bundles.

Adaptations:

- `apps/docs/src/renderer/office-host.ts`: explicit clipboard/language host.
- Renderer clipboard and language callbacks use OfficeHost, not preload globals.
- The Electron renderer entry explicitly injects its preload adapter, preserving
  the upstream desktop clipboard and language behavior.
- `packages/ui/src/office-render.ts`: document-only UI exports, leaving the
  upstream UI barrel intact. Rendering helpers import this entry explicitly.
- KompAgent's wrapper and planned typed iframe bridge belong in the parent
  repository; this source checkpoint does not certify Web Office acceptance.

To update upstream, fetch the `upstream` remote, review and merge a selected
upstream commit into this branch, resolve host-adapter conflicts, run the DOCX
round-trip/browser checks, commit here, then update the parent submodule pointer.
Do not change the parent pointer before the adapted child commit is available
from its configured origin.
