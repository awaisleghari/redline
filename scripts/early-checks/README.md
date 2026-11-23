# Early checks

One-off scripts behind the three "early checks" in `PLAN.md` (answers are in
its decisions log). They are evidence, not part of the product or the test
suite. Run them from the repository root; `OUT` is any empty scratch folder.

```bash
npx tsx scripts/early-checks/prepare.ts "$OUT"
```

writes every test pair, plus the redline the npm (WebAssembly) engine makes for it.

```bash
(cd scripts/early-checks/dotnet-compare && dotnet run -c Release -- "$OUT")
```

runs the .NET build of the same engine on the same pairs (needs the .NET 10 SDK).

```bash
npx tsx scripts/early-checks/compare-outputs.ts "$OUT"
```

compares the two sets of outputs part by part.

```bash
npx tsx scripts/early-checks/viewer.ts "$OUT"
```

converts redlines to HTML and reports how changes and redline styling are rendered.

Browser speed on the 100-page contract is measured by
`tests/e2e/large-document.spec.ts`, which prints the timings.
