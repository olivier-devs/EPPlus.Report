using OfficeOpenXml.ConditionalFormatting;

namespace EPPlus.Report.Model;

/// <summary>
///     Represents a conditional formatting rule extracted from a template block for later reconciliation.
/// </summary>
public class ConditionalFormattingRule
{
    /// <summary>
    ///     Gets or sets the cell address range to which the rule applies.
    /// </summary>
    public string Address { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the formula for the conditional formatting rule.
    /// </summary>
    public string Formula { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the second formula for conditional formatting rules that require two formulas.
    /// </summary>
    public string Formula2 { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the type of the conditional formatting rule.
    /// </summary>
    public eExcelConditionalFormattingRuleType Type { get; set; }

    /// <summary>
    ///     Gets or sets the priority of the conditional formatting rule.
    /// </summary>
    public int Priority { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether evaluation should stop if this rule evaluates to true.
    /// </summary>
    public bool StopIfTrue { get; set; }
}