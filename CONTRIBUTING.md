# Contributing

Issues and pull requests are welcome — from a typo in the docs to a new pipeline stage.

## 📐 Start with the architecture

Read **[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)** first: the bird's-eye view, the
code map, the tool pipelines and the invariants. Ten minutes there saves an hour in the
code — and most review feedback is about keeping those invariants intact.

## Build and test

```
build.cmd                  # dotnet SDK → single dist\iwoHelperDesktop.exe
build.cmd x86              # 32-bit build → dist\x86\iwoHelperDesktop.exe
tests\build_tests.cmd      # unit tests (no Office needed — what CI runs)
tests\build_tests.cmd x86  # the same tests as a 32-bit process (cache sizes branch on it)
tests\run_all.cmd          # full pyramid: build, units, GUI smoke, integration (needs Excel/Word)
```

Requirements: Windows 10/11 x64, the .NET SDK 8 or newer and Python 3.6+ (the commit hook
and the link check). Microsoft Office is needed only for the Excel and Word integration
scripts (`tests\verify*.ps1`) and for using the Merge Excel and PDF → Word tools themselves.

Some checks build real windows on an STA thread — they open, resize and close actual
forms, so a windowless session cannot run them.

## Ground rules

- **Language level is pinned**: C# 7.3 on .NET Framework 4.8 — no newer syntax.
- **Match the house style**: identifiers in English, comments in Russian (the project's
  working language), comment *why*, not *what*.
- **Logic goes into pure functions** with unit tests in `tests/UnitTests.cs`. Behaviour
  that needs Office gets a `tests\verify*.ps1` script instead.
- **Office COM rules are non-negotiable** — see
  [Office COM layer](docs/ARCHITECTURE.md#office-com-layer): release through
  `ComSafe`, never call a closed object, escape cell text via `CellText`.
- **No new runtime dependencies** unless embedded as a resource and compatible with
  GPL-3.0-only and the existing third-party licenses. Copyleft tools may run as separate
  processes, as Ghostscript does.
- **UI strings go through `Loc`** in both languages, generated documents (cover note,
  TOC, reports) intentionally stay Russian.

## Commits

The history holds commits by the maintainer only, with no `Co-authored-by` or other
attribution lines: GitHub builds the contributors list from them, and the only way to
correct it afterwards is rewriting history. Accepted pull requests are therefore landed by
the maintainer. `tools/check_commit_metadata.py` holds the rules; CI checks every commit and
tag pushed to `main` with it (not on pull requests: the maintainer lands accepted ones), and
the `commit-msg` hook applies the same rules before a commit exists — it drops attribution
lines from the message and refuses a foreign author. The hook needs Python 3.6+ on `PATH`;
without it every commit is refused. Enable it once per
clone:

```
git config core.hooksPath tools/githooks
```

## Pull requests

- Branch from `main`, keep the diff focused, add tests for new logic.
- CI must be green: build, unit tests, GUI smoke, dependency probes, installer check and
  README links (commit metadata is checked after landing).
- Describe the user-visible change in the PR text — it becomes a
  [CHANGELOG](docs/CHANGELOG.md) entry at release time.

Releases are cut by the maintainer (artifacts are signed locally) — see
[docs/RELEASING.md](docs/RELEASING.md).

## License

The project is licensed under the
[GNU General Public License v3.0 only](LICENSE).
By contributing, you agree that your contribution is licensed under the same terms,
unless a separate written agreement says otherwise.
