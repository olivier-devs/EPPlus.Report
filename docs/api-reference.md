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
| `IsNamedRangeLoop` | `bool` | Indicates whether the current context is inside a named range loop. |
| `CurrentIndex` | `int` | Zero-based index of the current item within the collection. |

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

## `ExpressionEvaluator` (Advanced)

The default implementation of `IExpressionEvaluator`. Resolves property paths via reflection with caching.

### Constructors

#### `ExpressionEvaluator()`

Initializes with built-in functions: `Upper`, `Lower`, `Trim`.

### Methods

#### `Evaluate(string expression, object context)`

Evaluates a property path against the context object.

#### `Evaluate(string expression, object context, string functionName)`

Evaluates a property path and applies the named function to the result.

#### `RegisterFunction(string name, Func<object, object> func)`

Registers a custom function by name.

#### `ApplyFunction(string functionName, object value)`

Applies a registered function to a value.

---

## Model Classes (Advanced)

### `TemplateNode` (abstract)

Base class for all template AST nodes.

| Property | Type | Description |
|----------|------|-------------|
| `Row` | `int` | Row number in the worksheet. |
| `Column` | `int` | Column number in the worksheet. |
| `RawContent` | `string` | Raw text content of the cell. |

### `ExpressionNode`

Represents a template expression `{{Property}}` or `{{Function(Property)}}`.

| Property | Type | Description |
|----------|------|-------------|
| `ExpressionPath` | `string` | Property path to evaluate (e.g. `"Object.Property"`). |
| `FunctionName` | `string` | Optional function name to apply (e.g. `"Upper"`). |

### `LoopNode`

Represents a `<<foreach Items>>` block.

| Property | Type | Description |
|----------|------|-------------|
| `CollectionName` | `string` | Name of the collection to iterate. |
| `Children` | `List<TemplateNode>` | Child nodes inside the loop block. |
| `EndRow` | `int` | Row where the loop block ends. |
| `ConditionalFormattingRules` | `List<ConditionalFormattingRule>` | CF rules associated with this block. |

### `IfNode`

Represents a `<<if Condition>>` block.

| Property | Type | Description |
|----------|------|-------------|
| `ConditionExpression` | `string` | Boolean expression to evaluate. |
| `Children` | `List<TemplateNode>` | Child nodes inside the conditional block. |
| `EndRow` | `int` | Row where the conditional block ends. |
| `ConditionalFormattingRules` | `List<ConditionalFormattingRule>` | CF rules associated with this block. |

### `GroupNode`

Represents a `<<group Items by Category>>` block. Inherits from `LoopNode`.

| Property | Type | Description |
|----------|------|-------------|
| `GroupByPaths` | `List<string>` | Property paths used to group items. |
| `Options` | `GroupOptions` | Options controlling group rendering. |
| `SubtotalTemplate` | `List<TemplateNode>` | Template nodes for subtotal rows. |

### `NamedRangeLoopNode`

Represents a loop derived from an Excel named range. Inherits from `LoopNode`.

| Property | Type | Description |
|----------|------|-------------|
| `RangeName` | `string` | Name of the Excel named range. |
| `IsHorizontal` | `bool` | Whether the loop iterates horizontally. |
| `ServiceRowCount` | `int` | Number of service rows at the end. |
| `ServiceTags` | `List<ServiceTag>` | Service tags (sum, count) in the service row. |
| `EndColumn` | `int` | Column where the named range ends. |
| `HeaderRowCount` | `int` | Number of header rows at the start. |
| `GroupByDefinitions` | `List<GroupByDefinition>` | Group-by definitions for named range grouping. |
| `RangeGroupOptions` | `GroupOptions` | Group options specific to the named range. |

### `AggregationNode`

Represents `<<sum Property>>` or `<<count Items>>`.

| Property | Type | Description |
|----------|------|-------------|
| `AggregationType` | `string` | Type of aggregation (`"sum"` or `"count"`). |
| `PropertyName` | `string` | Property or collection name to aggregate. |

### `GroupOptions`

Options for controlling group rendering behavior.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Collapse` | `bool` | `false` | Whether grouped rows should be collapsed. |
| `MergeLabels` | `MergeMode` | `None` | How to merge group label cells. |
| `PlaceToColumn` | `int` | `0` | Column index for group labels. |
| `WithHeader` | `bool` | `false` | Whether the group includes a header. |
| `DisableSubtotals` | `bool` | `false` | Whether to disable subtotal rows. |
| `DisableOutline` | `bool` | `false` | Whether to disable the group outline. |
| `PageBreaks` | `bool` | `false` | Whether to insert page breaks between groups. |
| `TotalLabel` | `string` | `"Total"` | Label text for total rows. |
| `GrandLabel` | `string` | `"Grand"` | Label text for grand total rows. |
| `SummaryAbove` | `bool` | `false` | Whether summary rows appear above detail rows. |
| `DisableGrandTotal` | `bool` | `false` | Whether to disable the grand total row. |
| `Descending` | `bool` | `false` | Whether to sort groups in descending order. |

### `MergeMode`

Defines how group labels should be merged.

| Value | Description |
|-------|-------------|
| `None` | Group labels are not merged. |
| `Merge1` | Merges group labels and clears duplicate cells. |
| `Merge2` | Merges group labels and clears duplicate cells (variant 2). |
| `Merge3` | Merges group labels without clearing duplicate cells. |

### `GroupByDefinition`

Defines a grouping criterion for named range loops.

| Property | Type | Description |
|----------|------|-------------|
| `PropertyPath` | `string` | Property path used to extract group keys. |
| `Column` | `int` | Column index where the group key is located. |
| `Descending` | `bool` | Whether the group should be sorted in descending order. |
| `Options` | `GroupOptions` | Options controlling rendering of this group. |

### `ServiceTag`

Represents a service tag in a named range (e.g. `<<sum>>`).

| Property | Type | Description |
|----------|------|-------------|
| `TagName` | `string` | Name of the tag (e.g. `"sum"` or `"count"`). |
| `Row` | `int` | Row where the tag is located. |
| `Column` | `int` | Column where the tag is located. |

### `ConditionalFormattingRule`

Represents a conditional formatting rule extracted from a template block.

| Property | Type | Description |
|----------|------|-------------|
| `Address` | `string` | Cell address range. |
| `Formula` | `string` | Primary formula. |
| `Formula2` | `string` | Secondary formula (for rules requiring two). |
| `Type` | `eExcelConditionalFormattingRuleType` | Type of conditional formatting rule. |
| `Priority` | `int` | Rule priority. |
| `StopIfTrue` | `bool` | Whether evaluation stops if this rule is true. |

### `PropertyNotFoundException`

Exception thrown when a property referenced in a template expression cannot be found. Inherits from `ArgumentException`.

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
- `asc` suffix for explicit ascending sort: `<<group Items by Category asc>>`
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

| Tag | Description | Excel Function |
|-----|-------------|-----------------|
| `<<sum>>` | Calculates and inserts the sum of values in the column | SUBTOTAL(9) - SUM |
| `<<count>>` | Calculates and inserts the count of non-empty values in the column | SUBTOTAL(3) - COUNTA |
| `<<avg>>` | Calculates and inserts the average of values in the column | SUBTOTAL(1) - AVERAGE |
| `<<counta>>` | Alias for <<count>>, counts non-empty values | SUBTOTAL(3) - COUNTA |
| `<<max>>` | Calculates and inserts the maximum value in the column | SUBTOTAL(4) - MAX |
| `<<min>>` | Calculates and inserts the minimum value in the column | SUBTOTAL(5) - MIN |
| `<<product>>` | Calculates and inserts the product of values in the column | SUBTOTAL(6) - PRODUCT |
| `<<stddev>>` | Calculates and inserts the sample standard deviation | SUBTOTAL(7) - STDEV |
| `<<stddevp>>` | Calculates and inserts the population standard deviation | SUBTOTAL(8) - STDEVP |
| `<<var>>` | Calculates and inserts the sample variance | SUBTOTAL(10) - VAR |
| `<<varp>>` | Calculates and inserts the population variance | SUBTOTAL(11) - VARP |

All service tags generate dynamic `SUBTOTAL()` formulas that recalculate automatically when users modify data in Excel.

---

## Conditional Formatting Style Preservation

All visual styles of conditional formatting rules are now fully cloned during rendering:

- **Fill, Font, Border styles** — preserved for Expression, CellIs, GreaterThan, LessThan, Equal, and similar rule types
- **Color Scales** — TwoColorScale and ThreeColorScale with all color stops and value types
- **Data Bars** — Color, min/max bounds, axis position, direction, border
- **Icon Sets** — 3/4/5 icon sets with ShowValue, Reverse, and individual criteria
- **Unsupported types** — gracefully fallback to v1 behavior (red placeholder)

---

## Thread Safety

`TemplateEngine` is **not thread-safe**. Create one instance per generation operation.
