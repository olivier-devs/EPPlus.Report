using System.Collections.Generic;
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

    /// <summary>
    ///     Gets or sets the fill style, including pattern type and colors.
    ///     Null when the rule type does not support fill styles.
    /// </summary>
    public CfFillStyle FillStyle { get; set; }

    /// <summary>
    ///     Gets or sets the font style, including bold, italic, color, etc.
    ///     Null when the rule type does not support font styles.
    /// </summary>
    public CfFontStyle FontStyle { get; set; }

    /// <summary>
    ///     Gets or sets the border style for each edge.
    ///     Null when the rule type does not support border styles.
    /// </summary>
    public CfBorderStyle BorderStyle { get; set; }

    /// <summary>
    ///     Gets or sets the color scale stops for TwoColorScale / ThreeColorScale rules.
    ///     Null when the rule type is not a color scale.
    /// </summary>
    public List<CfColorStop> ColorScaleStops { get; set; }

    /// <summary>
    ///     Gets or sets the data bar settings for DataBar rules.
    ///     Null when the rule type is not a data bar.
    /// </summary>
    public CfDataBarSettings DataBarSettings { get; set; }

    /// <summary>
    ///     Gets or sets the icon set settings for IconSet rules.
    ///     Null when the rule type is not an icon set.
    /// </summary>
    public CfIconSetSettings IconSetSettings { get; set; }
}