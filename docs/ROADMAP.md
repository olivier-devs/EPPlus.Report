# EPPlus.Report - Roadmap

## Status: 2026-04-25

### Completed

#### API Style ClosedXML.Report (DONE)
- `TemplateEngine` refactored with `AddVariable(object)` and `AddVariable(string, object)`
- `Generate(GenerateOptions options = null)` — simplified overload with optional parameter
- `Save()` / `SaveAs(string)` / `SaveAs(FileInfo)` / `SaveAs(Stream)` with `SaveOptions`
- Support for formula evaluation via `GenerateOptions.EvaluateFormulas` and `SaveOptions.EvaluateFormulasBeforeSave`
- Error collection during parsing (`TemplateErrors`, `TemplateError`, `ErrorType`)
- Named variable resolution in renderer via `RenderContext.Variables`
- Stream-based template loading

#### Named Range Flat Tables (DONE)
- Named Range scanning via `NamedRangeScanner`
- `NamedRangeLoopNode` with service row support
- Special variables: `item`, `item.Property`, `index`, `items`
- Service tags: `<<sum>>` (SUBTOTAL formula), `<<count>>`
- Coexists with existing `<<foreach>>` syntax
- Priority: explicit `<<foreach>>` > Named Range

#### Documentation (DONE)
- `docs/api-reference.md` — Full API reference (English)
- `docs/api-reference.fr.md` — Référence API complète (Français)
- `docs/usage-guide.md` — Usage guide with examples (English)
- `docs/usage-guide.fr.md` — Guide d'utilisation avec exemples (Français)
- `docs/ROADMAP.md` — Development roadmap
- `AGENTS.md` — Updated with new API examples

#### Advanced Error Reporting (DONE)
- `TemplateError` enriched with `WorksheetName`, `Row`, `Column`, `Expression`, `Location`
- `ErrorType.Rendering` added for Excel operation errors
- `TemplateGenerateResult` exposes `ParsingErrors` and `RenderingErrors`
- `TemplateRenderer` collects `ArgumentException` and `NullReferenceException` gracefully
- Backward compatibility preserved (exceptions propagate when no collector provided)
- **`TemplateGenerateResult.Warnings`** for non-blocking issues (e.g. missing properties) — rendering continues

### In Progress

None currently.

### Planned

- [x] Grouping directive: `<<group Items by Category>>`
  - Explicit block syntax: `<<group Items by Category>>` ... `<</group>>`
  - Named Range service row syntax: `<<group Category>>`
  - Subtotals (`<<sum>>`, `<<count>>`) and Grand Total per group
  - MergeLabels, WithHeader, DisableSubtotals options
- [x] Advanced error reporting (line numbers, worksheet names in errors)
- [ ] `SaveOptions.Password` for workbook encryption
- [x] `TemplateGenerateResult.Warnings` for non-blocking issues
- [x] Performance benchmarks (>10k rows in <30s)
- [x] Documentation: full API reference and usage guide

### Completed (2026-04-26)

#### Custom Functions in Expressions (DONE)
- Syntax: `{{Upper(Name)}}`, `{{Trim(Address.City)}}`
- Built-in functions: `Upper`, `Lower`, `Trim`
- User-registered functions via `TemplateEngine.RegisterFunction(name, func)`
- Full integration: parser → AST (`ExpressionNode.FunctionName`) → renderer → evaluator
- 8 integration tests in `CustomFunctionTests.cs`

#### Conditional Formatting Preservation (DONE)
- Parser scans `worksheet.ConditionalFormatting` and associates rules with block nodes
- Renderer reconciles CF after each block render (remove old rules, re-apply to final range)
- Supports loops, if blocks, groups, and named range loops
- 5 integration tests in `ConditionalFormattingTests.cs`
- **Limitation v1:** Style recreated as placeholder; full style clone planned for v2

#### ExcelTable Preservation (DONE)
- `RowOperationTracker` records all `InsertRow`/`DeleteRow` during rendering
- `ExcelTableAdjuster` adjusts `TableRange` post-render
- Discovered: EPPlus 7.x auto-adjusts table addresses on row insert/delete, so adjuster is minimal
- 5 integration tests in `ExcelTableTests.cs`

#### Warnings in TemplateGenerateResult (DONE)
- `ErrorType.Warning` added for non-blocking issues
- `PropertyNotFoundException` introduced for missing property evaluation failures
- `TemplateRenderer` catches missing properties and routes them to `Warnings` instead of `RenderingErrors`
- `TemplateGenerateResult` exposes `Warnings` and `HasWarnings`
- Backward compatibility: when no warning collector provided, falls back to `RenderingErrors` or throws
- 3 tests updated: `SimpleExpressionTests`, `RendererTests`, `TemplateEngineTests`

#### Performance Benchmarks (DONE)
- BenchmarkDotNet project `EPPlus.Report.Benchmarks` created
- Benchmarks cover: simple expressions, loops, groupings, named range loops
- **Results (Intel Core i5-4460, .NET 8.0):**
  - `LoopRender` (10k rows): ~70ms
  - `LoopRender` (20k rows): ~152ms
  - `NamedRangeLoopRender` (10k rows): ~20ms
  - `GroupedLoopRender` (10k rows): ~211ms
- **Target achieved:** >10,000 rows in <30s (actual: <0.3s for simple loops)
- Note: `TemplateEngine` with `<<foreach>>` blocks has a parsing bug when loaded from file/stream; benchmarks use direct `TemplateParser`/`TemplateRenderer` for loop scenarios

### Security Review (2026-05-03)

#### Security Audit Complete
- **Injection**: SAFE - Expression evaluation uses reflection with runtime type validation, no code execution
- **XSS/CSRF**: NOT APPLICABLE - Library generates Excel files, not HTML
- **Authentication**: NOT APPLICABLE - Server-side templating library
- **Secrets Management**: SAFE - No hardcoded secrets, no sensitive data logging
- **Input Validation**: ADEQUATE - Null checks, empty string validation, regex patterns safe (no ReDoS)
- **Logging**: SAFE - Errors collected in `TemplateErrors`, no sensitive data exposure

**Verdict**: APPROVED - No critical vulnerabilities found

### Backlog

- [ ] Pivot tables support (beyond EPPlus auto-adjustment)
- [ ] Full CF style cloning (colors, data bars, icon sets)
- [ ] Async API variants
- [ ] `SaveOptions.Password` for workbook encryption
- [x] `TemplateGenerateResult.Warnings` for non-blocking issues
- [x] Performance benchmarks (>10k rows in <30s)
