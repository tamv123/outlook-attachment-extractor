# Contributing

Thanks for helping improve the Outlook Attachment Extractor!

## Branching & pull requests

- **Never commit directly to `main`.** `main` stays stable and releasable.
- Do your work on a **`dev`** branch (or a `feature/<name>` branch off `dev`).
- When the change is finished **and confirmed working**, open a **Pull Request**
  into `main` for review.
- `main` is only updated through merged PRs.

## Versioning (Semantic Versioning)

This project follows [SemVer](https://semver.org/): **MAJOR.MINOR.PATCH**.

- **MAJOR** — incompatible changes to parameters/behavior users rely on.
- **MINOR** — new, backwards-compatible features (e.g. a new option).
- **PATCH** — backwards-compatible bug fixes.

When you change the version, update **all** of these together in the same PR:

1. `$ScriptVersion` in **`Extract-Attachments.ps1`**
2. `$ScriptVersion` in **`Run-Extractor.ps1`**
3. The **version badge** at the top of `README.md`
4. A new entry in the **Changelog** section of `README.md`

After the PR merges to `main`, tag the release:

```bash
git checkout main && git pull
git tag -a v1.2.0 -m "v1.2.0"
git push origin v1.2.0
```

(Optionally publish a GitHub Release from that tag.)

## Code style — keep it ASCII

Windows PowerShell 5.1 reads a BOM-less `.ps1` using the system ANSI code page.
Non-ASCII characters (em-dashes `—`, smart quotes, emoji) then get mis-decoded
and can **corrupt string parsing**, breaking the whole script.

- Use plain ASCII in `.ps1` code: `-` instead of `—`, straight quotes `'` `"`.
- For symbols in generated email HTML, use **HTML entities** (e.g. `&#128206;`),
  not literal emoji.
- Quick check before committing:
  ```powershell
  $e=$null; [System.Management.Automation.Language.Parser]::ParseFile(
    (Resolve-Path .\Extract-Attachments.ps1), [ref]$null, [ref]$e); $e
  ```
  No output = clean.

## Outlook COM gotcha — avoid `GetItemFromID` for bulk work

`GetItemFromID` triggers a server fetch on uncached or stale EntryIDs and can
hang indefinitely. For anything that revisits multiple emails, iterate
`$folder.Items` (cache-resident) instead. See the README "Developer Note".
