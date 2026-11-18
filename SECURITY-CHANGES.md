# Security maintenance

## Unreleased

- The Snake plugin accepts score and game-over messages only from the registered
  iframe at the document origin. A browser regression checks forged parent and
  sibling sources, a foreign origin, and legitimate game messages.
- Update the build and release toolchain and pin patched transitive dependencies.
  Keep legacy CommonJS major APIs where callers require them. Vite remains on
  the patched 6.4 branch for VitePress compatibility.
- Resolve the upstream editor declaration's private Print alias explicitly for
  DOCX typechecking; plugin source and dependency declarations remain checked.
- CI runs the audit, typecheck, lint, browser regression, library/demo builds and
  documentation build on Node 24 and 26. Secret scanning and push protection are
  complemented by Dependabot and CodeQL monitoring.
- 2026-11-18: `pnpm audit` ignores exactly GHSA-fc5q-39xc-vjpf (picomatch
  <=4.0.7 ReDoS through repeated extglobs). picomatch is reachable only from
  dev-only tooling (`@typescript-eslint` via globby, fast-glob and micromatch;
  `lerna` via tinyglobby), never from a production dependency
  (`pnpm audit --prod` is clean), and no fixed release exists. Remove the
  `auditConfig.ignoreGhsas` entry in `pnpm-workspace.yaml` once picomatch
  publishes a fix.

These are changes to this maintained fork. They do not publish packages under
upstream's npm scope or submit changes to the upstream project.
