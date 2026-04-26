# EPPlus.Report - API Reference

Complete reference of the public API.

---

## `TemplateEngine`

The main entry point for template-based Excel generation. Inspired by ClosedXML.Report.

### Constructors

#### `TemplateEngine(string templatePath)`

Creates a new engine from an Excel template file on disk.

**Parameters:**
- `templatePath` (`string`): Path to the `.xlsx` template file.

**Exceptions:**
- `ArgumentNullException`: If `templatePath` is null.

#### `TemplateEngine(Stream stream)`

Creates a new engine from a stream containing an Excel workbook.

**Parameters:**
- `stream` (`Stream`): Readable stream containing a valid `.xlsx` file.

**Exceptions:**
- `ArgumentNullException`: If `stream` is null.

---

### Methods

#### `AddVariable(object value)`

Sets the root context object. Used when an expression does not match any named variable.

**Parameters:**
- `value` (`object`): The root data object.

#### `AddVariable(string name, object value)`

Adds a named variable accessible in templates via `{{Name}}`.

**Parameters:**
- `name` (`string`): Variable name used in templates.
- `value` (`object`): Variable value.

**Exceptions:**
- `ArgumentException`: If `name` is null, empty, or whitespace.

**Resolution priority:** Named variables are resolved first. If no named variable matches an expression, the root context (`AddVariable(object)`) is used as fallback.

#### `Generate()`

Parses all worksheets, renders the template, and returns the result.

**Returns:** `TemplateGenerateResult` containing any parsing or rendering errors.

**Behavior:**
- Modifies the internal `ExcelPackage` in memory.
- Does not write to disk.
- Formulas are not evaluated (leave that to Excel at open time).

#### `Generate(GenerateOptions options)`

Renders the template with additional options.

**Parameters:**
- `options` (`GenerateOptions`): Generation options.

**Returns:** `TemplateGenerateResult`

#### `RegisterFunction(string name, Func<object, object> func)`

Registers a custom function available in template expressions.

**Parameters:**
- `name` (`string`): Function name used in templates.
- `func` (`Func<object, object>`): Function implementation.

**Exceptions:**
- `ArgumentException`: If `name` is null, empty, or whitespace.
- `ArgumentNullException`: If `func` is null.

**Example:**
```csharp
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
// Template: {{Double(Name)}}
```

**Built-in functions:** `Upper`, `Lower`, `Trim`.

---

### Save Methods

#### `Save()`

Overwrites the original template file.

**Exceptions:**
- `InvalidOperationException`: If the engine was created from a stream (no original path exists).

#### `SaveAs(string path)`

Saves to a new file path.

**Parameters:**
- `path` (`string`): Output file path.

#### `SaveAs(FileInfo fileInfo)`

Saves using a `FileInfo`.

**Parameters:**
- `fileInfo` (`FileInfo`): Target file info.

#### `SaveAs(Stream stream)`

Writes the workbook to a stream.

**Parameters:**
- `stream` (`Stream`): Writable stream.

#### `SaveAs(string path, SaveOptions saveOptions)`

Saves to a path with options.

**Parameters:**
- `path` (`string`): Output file path.
- `saveOptions` (`SaveOptions`): Save options (e.g. formula evaluation).

#### `SaveAs(FileInfo fileInfo, SaveOptions saveOptions)`

Saves using `FileInfo` with options.

#### `SaveAs(Stream stream, SaveOptions saveOptions)`

Writes to stream with options.

---

## `TemplateGenerateResult`

Result object returned by `TemplateEngine.Generate()`.

### Properties

#### `HasErrors` (get)

`bool` - `true` if at least one error was collected during parsing or rendering.

#### `HasWarnings` (get)

`bool` - `true` if at least one non-blocking warning was collected during rendering.

#### `ParsingErrors` (get)

`TemplateErrors` - Collection of parsing errors (e.g. unclosed blocks). Empty if no errors occurred.

#### `RenderingErrors` (get)

`TemplateErrors` - Collection of rendering errors (e.g. formula evaluation failures). Empty if no errors occurred.

#### `Warnings` (get)

`TemplateErrors` - Collection of non-blocking warnings (e.g. missing properties). Rendering continues. Empty if no warnings occurred.

---

## `TemplateErrors`

Inherits `List<TemplateError>`. Collection of template errors.

---

## `TemplateError`

Represents a single error encountered during template processing.

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `Message` | `string` | Human-readable error description. |
| `CellAddress` | `string` | Excel cell address (e.g. `"A5"`). |
| `WorksheetName` | `string` | Name of the worksheet where the error occurred. |
| `Row` | `int` | Row number (1-based). |
| `Column` | `int` | Column number (1-based). |
| `Expression` | `string` | The template expression that caused the error. |
| `Location` | `string` | Combined location: `WorksheetName!CellAddress`. |
| `Type` | `ErrorType` | Error category: `Parsing`, `Evaluation`, `Rendering`, or `Warning`. |

---

## `ErrorType`

```csharp
public enum ErrorType
{
    Parsing,     // Template syntax errors (unclosed blocks, etc.)
    Evaluation,  // Expression evaluation errors (fatal)
    Rendering,   // Excel operation errors during rendering
    Warning      // Non-blocking issues (e.g. missing properties)
}
```

---

## `GenerateOptions`

Options for `TemplateEngine.Generate(GenerateOptions)`.

### Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `EvaluateFormulas` | `bool` | `false` | If `true`, evaluates all Excel formulas after rendering via `package.Workbook.Calculate()`. |

---

## `SaveOptions`

Options for save operations.

### Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `EvaluateFormulasBeforeSave` | `bool` | `false` | If `true`, evaluates formulas just before writing to disk/stream. |

> **Note:** If `GenerateOptions.EvaluateFormulas` was already `true`, formulas are not re-evaluated during save.

---

## `RenderContext` (Advanced)

Passed internally to the renderer. Available for advanced usage.

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `Current` | `object` | Current data context (root or loop item). |
| `Variables` | `Dictionary<string, object>` | Named variables dictionary. |
| `CurrentCollection` | `IEnumerable` | Current collection context for aggregations. |

---

## Interfaces (Extension Points)

### `ITemplateParser`

```csharp
public interface ITemplateParser
{
    Template Parse(ExcelWorksheet worksheet, TemplateErrors errors);
}
```

### `ITemplateRenderer`

```csharp
public interface ITemplateRenderer
{
    void Render(Template template, RenderContext context, ExcelWorksheet worksheet);
}
```

### `IExpressionEvaluator`

```csharp
public interface IExpressionEvaluator
{
    object Evaluate(string expression, object context);
}
```

---

## Template Directives

| Directive | Syntax | Description |
|-----------|--------|-------------|
| Expression | `{{Property}}` | Replaced with property value |
| Nested expression | `{{Object.Property}}` | Replaced with nested property value |
| Loop | `<<foreach Items>>` | Repeats block for each item |
| Loop end | `<</foreach>>` | Closes loop block |
| Condition | `<<if Condition>>` | Shows block if condition is true |
| Condition end | `<</if>>` | Closes condition block |
| Sum | `<<sum Property>>` | Calculates sum of property across collection |
| Count | `<<count Items>>` | Counts items in collection |
| Group | `<<group Items by Category>>` | Groups collection by property with subtotals |
| Group end | `<</group>>` | Closes group block |

---

## Custom Functions

Functions transform expression values at render time.

### Built-in Functions

| Function | Example | Result |
|----------|---------|--------|
| `Upper` | `{{Upper(Name)}}` | Converts to uppercase |
| `Lower` | `{{Lower(Name)}}` | Converts to lowercase |
| `Trim` | `{{Trim(Name)}}` | Removes leading/trailing whitespace |

### User-registered Functions

Register via `TemplateEngine.RegisterFunction(name, func)`:

```csharp
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
// Template: {{Double(Name)}}
```

Functions are applied after property resolution. If the property value is `null`, the function receives `null`.

---

## Grouping

Group a collection by one or more properties, with optional subtotals and grand totals.

### Explicit Block Syntax

```
<<group Items by Category>>
{{Category}}
{{Name}} | {{Price}}
<</group>>
```

**Options:**
- `desc` suffix for descending sort: `<<group Items by Category desc>>`
- `MergeLabels` merges group key cells vertically
- Subtotal template row (last row with only aggregation nodes) is auto-detected
- Grand total is rendered after all groups (unless disabled)

### Named Range Grouping

In Named Range flat tables, use the service row syntax:

```
<<group Category>>
```

This groups the named range loop by the specified property and renders subtotal rows per group.

---

## Named Range Flat Tables

An alternative to `<<foreach>>` loops using Excel Named Ranges.

### How it works

1. Create a **Named Range** in Excel with the same name as your collection variable.
2. Inside the named range, use `{{item.Property}}` to reference element properties.
3. The last row of the named range is the **service row** for tags like `<<sum>>`.

### Requirements

- Named range must be rectangular and continuous (no gaps).
- Vertical tables: at least 2 rows and 2 columns.
- The last row is reserved for service tags.

### Special variables

| Variable | Resolution |
|----------|------------|
| `item` | Current element of the collection |
| `item.Property` | Property of the current element |
| `index` | 0-based index of the current element |
| `items` | The entire collection |

### Service row tags

| Tag | Description |
|-----|-------------|
| `<<sum>>` | Inserts a `=SUBTOTAL(9, ...)` formula for the column |
| `<<count>>` | Inserts a `=SUBTOTAL(3, ...)` formula for the column |

### Example

**Template Excel:** Named Range `Orders` covering `A1:B3`

| | A | B |
|---|---|---|
| 1 | Order No | Amount |
| 2 | {{item.OrderNo}} | {{item.Amount}} |
| 3 | <<sum>> | |

**Code:**
```csharp
var engine = new TemplateEngine("template.xlsx");
engine.AddVariable("Orders", new[]
{
    new { OrderNo = 100, Amount = 50m },
    new { OrderNo = 101, Amount = 75m }
});
var result = engine.Generate();
engine.SaveAs("output.xlsx");
```

**Result:** Rows A2:B2 are duplicated for each order, with `=SUBTOTAL(9, ...)` in A3.

---

## Thread Safety

`TemplateEngine` is **not thread-safe**. Create one instance per generation operation.
