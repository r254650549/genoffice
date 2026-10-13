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
the planned Web/runtime bundles. The complete Web host is still pending.

Adaptations:

- `apps/docs/src/renderer/office-host.ts`: explicit clipboard/language host.
- Renderer clipboard and language callbacks use OfficeHost, not preload globals.
- The Electron renderer entry explicitly injects its preload adapter, preserving
  the upstream desktop clipboard and language behavior.
- The document-only UI barrel from the rejected reduced-editor approach has
  been removed; rendering helpers use the original complete upstream UI entry.
- The parent repository will host the complete renderer in an isolated iframe;
  this source checkpoint does not certify Web Office acceptance.

To update upstream, fetch the `upstream` remote, review and merge a selected
upstream commit into this branch, resolve host-adapter conflicts, run the DOCX
round-trip/browser checks, commit here, then update the parent submodule pointer.
Do not change the parent pointer before the adapted child commit is available
from its configured origin.
