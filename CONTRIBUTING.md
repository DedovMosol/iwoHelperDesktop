# Contributing

Issues and pull requests are welcome — from a typo in the docs to a new pipeline stage.

## 📐 Start with the architecture

Read **[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)** first: the bird's-eye view, the
code map, the tool pipelines and the invariants. Ten minutes there saves an hour in the
code — and most review feedback is about keeping those invariants intact.

## Build and test

```
build.cmd                 # dotnet SDK → single dist\iwoHelperDesktop.exe
tests\build_tests.cmd     # unit tests (no Office needed — what CI runs)
tests\run_all.cmd         # full pyramid: build, units, GUI smoke, integration (needs Excel/Word)
```

Requirements: Windows 10/11 x64 and the `dotnet` SDK. Microsoft Office is needed only
for the integration scripts (`tests\verify*.ps1`) and for using the Excel Digest and
PDF → Word tools themselves.

## Ground rules

- **Language level is pinned**: C# 7.3 on .NET Framework 4.8 — no newer syntax.
- **Match the house style**: identifiers in English; comments in Russian (the project's
  working language); comment *why*, not *what*.
- **Logic goes into pure functions** with unit tests in `tests/UnitTests.cs`. Behaviour
  that needs Office gets a `tests\verify*.ps1` script instead.
- **Office COM rules are non-negotiable** — see
  [Office COM layer](docs/ARCHITECTURE.md#office-com-layer): release through
  `ComSafe`, never call a closed object, escape cell text via `CellText`.
- **No new runtime dependencies** unless embedded as a resource and MIT-compatible;
  copyleft tools may only run as separate processes (like Ghostscript does).
- **UI strings go through `Loc`** in both languages; generated documents (cover note,
  TOC, reports) intentionally stay Russian.

## Pull requests

- Branch from `main`, keep the diff focused, add tests for new logic.
- CI must be green: build, unit tests, GUI smoke, dependency probes, installer check.
- Describe the user-visible change in the PR text — it becomes a
  [CHANGELOG](docs/CHANGELOG.md) entry at release time.

Releases are cut by the maintainer (artifacts are signed locally) — see
[docs/RELEASING.md](docs/RELEASING.md).

## License

MIT. By contributing you agree that your contribution is licensed under the
[same terms](LICENSE).
