using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;

namespace EPPlus.Report.Model;

/// <summary>
///     Fill style for a conditional formatting rule.
/// </summary>
public class CfFillStyle
{
    public ExcelFillStyle PatternType { get; set; }
    public Color BackgroundColor { get; set; }
    public Color ForegroundColor { get; set; }
}

/// <summary>
///     Font style for a conditional formatting rule.
/// </summary>
public class CfFontStyle
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public Color Color { get; set; }
    public float Size { get; set; }
    public string Name { get; set; }
}

/// <summary>
///     Border style for a conditional formatting rule.
/// </summary>
public class CfBorderStyle
{
    public CfBorderEdge Top { get; set; }
    public CfBorderEdge Bottom { get; set; }
    public CfBorderEdge Left { get; set; }
    public CfBorderEdge Right { get; set; }
}

/// <summary>
///     A single edge of a border.
/// </summary>
public class CfBorderEdge
{
    public ExcelBorderStyle Style { get; set; }
    public Color Color { get; set; }
}

/// <summary>
///     A color stop in a color scale (min, mid, max point).
/// </summary>
public class CfColorStop
{
    public Color Color { get; set; }
    public eExcelConditionalFormattingValueObjectType Type { get; set; }
    public double? Value { get; set; }
}

/// <summary>
///     Settings for a DataBar conditional formatting rule.
/// </summary>
public class CfDataBarSettings
{
    public Color Color { get; set; }
    public CfValueBound MinValue { get; set; }
    public CfValueBound MaxValue { get; set; }
    public bool ShowValue { get; set; }
    public Color? AxisColor { get; set; }
    /// <summary>Axis position stored as int since EPPlus enum availability varies by version.</summary>
    public int? AxisPosition { get; set; }
    public Color? BorderColor { get; set; }
    /// <summary>DataBar direction stored as int since EPPlus enum availability varies by version.</summary>
    public int? Direction { get; set; }
}

/// <summary>
///     A min/max bound value for a DataBar.
/// </summary>
public class CfValueBound
{
    public eExcelConditionalFormattingValueObjectType Type { get; set; }
    public double? Value { get; set; }
}

/// <summary>
///     Settings for an IconSet conditional formatting rule.
/// </summary>
public class CfIconSetSettings
{
    /// <summary>Icon set type stored as int since EPPlus uses separate enums per icon count (3/4/5).</summary>
    public int IconSetType { get; set; }
    public bool ShowValue { get; set; }
    public bool Reverse { get; set; }
    public List<CfIconCriterion> Criteria { get; set; }
}

/// <summary>
///     A single criterion in an IconSet (e.g., >= 33%).
/// </summary>
public class CfIconCriterion
{
    public eExcelConditionalFormattingValueObjectType Type { get; set; }
    public double? Value { get; set; }
    /// <summary>Operator stored as int since EPPlus enum availability varies by version.</summary>
    public int? Operator { get; set; }
}