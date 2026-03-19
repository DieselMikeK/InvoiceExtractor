## Update Release Flow

### One-time setup

1. In GitHub repo settings, enable Actions and set workflow permissions to `Read and write`.
2. If `main` is branch-protected, allow GitHub Actions to push to it or this workflow will publish the exe but fail to update [docs/release.json](./docs/release.json).
3. On this computer, sign into GitHub CLI once with `gh auth login`.

### Normal release flow

1. Push the code you want released to `main`.
2. Change [VERSION](./VERSION) before that push if you want a new release number.
3. Run `.\publish_release.ps1` from the repo root.
4. Optionally pass notes with `.\publish_release.ps1 -Notes "What changed"`.

That command dispatches `.github/workflows/release.yml`, which:

- builds `InvoiceExtractorUpdater.exe` and `InvoiceExtractor.exe`
- creates GitHub release `v<version>`
- uploads `dist\InvoiceExtractor.exe`
- updates [docs/release.json](./docs/release.json) with the download URL, SHA-256, notes, and publish time
- pushes the manifest commit so client apps see the update button

Handoff layout after `.\build_release.ps1`:
- root: `InvoiceExtractor.exe`
- `app\required\...`
- `app\update\InvoiceExtractorUpdater.exe`

Normal pushes do not trigger client updates. The only client-visible release signal is [docs/release.json](./docs/release.json), and that file is only updated by the release workflow. Clients only see the `Update` button when the remote manifest version in [docs/release.json](./docs/release.json) is newer than their local [VERSION](./VERSION).
