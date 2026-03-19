## Update Release Flow

### One-time setup

1. In GitHub repo settings, enable Actions and set workflow permissions to `Read and write`.
2. If `main` is branch-protected, allow GitHub Actions to push to it or this workflow will publish the exe but fail to update [docs/release.json](./docs/release.json).

### Normal release flow

1. Push the code you want released to `main`.
2. Change [VERSION](./VERSION) before that push if you want a new release number.
3. Run `.\publish_release.ps1` from the repo root.
4. Optionally pass notes with `.\publish_release.ps1 -Notes "What changed"`.
5. Optionally let the script bump the version for the release commit with `.\publish_release.ps1 -Version 1.2.1 -Notes "What changed"`.

That command updates [release_request.json](./release_request.json), commits it, and pushes it to `main`. The `.github/workflows/release.yml` workflow watches that file and then:

- builds `InvoiceExtractorUpdater.exe` and `InvoiceExtractor.exe`
- creates GitHub release `v<version>`
- uploads `dist\InvoiceExtractor.exe` plus curated app payload files such as `app\vendors.csv`
- updates [docs/release.json](./docs/release.json) with the primary exe URL, SHA-256, notes, publish time, and the file list for the updater
- pushes the manifest commit so client apps see the update button

Handoff layout after `.\build_release.ps1`:
- root: `InvoiceExtractor.exe`
- `app\required\...`
- `app\update\InvoiceExtractorUpdater.exe`

Normal pushes do not trigger client updates. There are two separate files involved:

- [release_request.json](./release_request.json): the release trigger file that tells GitHub Actions to publish a release
- [docs/release.json](./docs/release.json): the client-visible manifest that makes installed apps show the `Update` button

The updater only downloads the curated file list in [docs/release.json](./docs/release.json). User-specific folders such as `required\`, and development folders such as `build\` and `training\`, are intentionally excluded from release updates.

Clients only see the `Update` button when the remote manifest version in [docs/release.json](./docs/release.json) is newer than their local [VERSION](./VERSION).
