using System;
using System.Collections.Generic;
using System.Drawing;
using EPPlus.Report.Model;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Extracts and applies conditional formatting style properties between EPPlus rules and <see cref="ConditionalFormattingRule" /> DTOs.
/// </summary>
public static class ConditionalFormattingCloner
{
    /// <summary>
    ///     Extracts style properties from an EPPlus conditional formatting rule into the specified DTO.
    /// </summary>
    /// <param name="source">The EPPlus conditional formatting rule to extract from.</param>
    /// <param name="target">The DTO to populate with extracted properties.</param>
    public static void Extract(IExcelConditionalFormattingRule source, ConditionalFormattingRule target)
    {
        try
        {
            switch (source.Type)
            {
                case eExcelConditionalFormattingRuleType.Expression:
                case eExcelConditionalFormattingRuleType.GreaterThan:
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                case eExcelConditionalFormattingRuleType.LessThan:
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                case eExcelConditionalFormattingRuleType.Equal:
                case eExcelConditionalFormattingRuleType.NotEqual:
                case eExcelConditionalFormattingRuleType.ContainsText:
                case eExcelConditionalFormattingRuleType.NotContainsText:
                case eExcelConditionalFormattingRuleType.BeginsWith:
                case eExcelConditionalFormattingRuleType.EndsWith:
                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                case eExcelConditionalFormattingRuleType.ContainsErrors:
                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                case eExcelConditionalFormattingRuleType.DuplicateValues:
                case eExcelConditionalFormattingRuleType.UniqueValues:
                case eExcelConditionalFormattingRuleType.Last7Days:
                case eExcelConditionalFormattingRuleType.LastMonth:
                case eExcelConditionalFormattingRuleType.LastWeek:
                case eExcelConditionalFormattingRuleType.NextMonth:
                case eExcelConditionalFormattingRuleType.NextWeek:
                case eExcelConditionalFormattingRuleType.ThisMonth:
                case eExcelConditionalFormattingRuleType.ThisWeek:
                case eExcelConditionalFormattingRuleType.Today:
                case eExcelConditionalFormattingRuleType.Tomorrow:
                case eExcelConditionalFormattingRuleType.Yesterday:
                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.TopPercent:
                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                case eExcelConditionalFormattingRuleType.Between:
                case eExcelConditionalFormattingRuleType.NotBetween:
                case eExcelConditionalFormattingRuleType.NotContains:
                    ExtractCellStyle(source, target);
                    break;

                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    ExtractTwoColorScale(source, target);
                    break;

                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    ExtractThreeColorScale(source, target);
                    break;

                case eExcelConditionalFormattingRuleType.DataBar:
                    ExtractDataBar(source, target);
                    break;

                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                case eExcelConditionalFormattingRuleType.FourIconSet:
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    ExtractIconSet(source, target);
                    break;

                default:
                    // Leave style properties null (v1 fallback)
                    break;
            }
        }
        catch
        {
            // Silently ignore extraction errors
        }
    }

    /// <summary>
    ///     Applies stored style properties from a DTO to a new EPPlus conditional formatting rule.
    /// </summary>
    /// <param name="source">The DTO containing stored properties.</param>
    /// <param name="worksheet">The worksheet to add the rule to.</param>
    /// <param name="newAddress">The cell address range for the new rule.</param>
    /// <returns>The newly created EPPlus conditional formatting rule.</returns>
    public static IExcelConditionalFormattingRule Apply(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        try
        {
            return source.Type switch
            {
                eExcelConditionalFormattingRuleType.Expression => ApplyExpression(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.GreaterThan => ApplyGreaterThan(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.GreaterThanOrEqual => ApplyGreaterThanOrEqual(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.LessThan => ApplyLessThan(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.LessThanOrEqual => ApplyLessThanOrEqual(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Equal => ApplyEqual(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NotEqual => ApplyNotEqual(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ContainsText => ApplyContainsText(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NotContainsText => ApplyNotContainsText(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.BeginsWith => ApplyBeginsWith(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.EndsWith => ApplyEndsWith(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ContainsBlanks => ApplyContainsBlanks(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NotContainsBlanks => ApplyNotContainsBlanks(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ContainsErrors => ApplyContainsErrors(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NotContainsErrors => ApplyNotContainsErrors(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.DuplicateValues => ApplyDuplicateValues(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.UniqueValues => ApplyUniqueValues(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Last7Days => ApplyLast7Days(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.LastMonth => ApplyLastMonth(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.LastWeek => ApplyLastWeek(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NextMonth => ApplyNextMonth(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NextWeek => ApplyNextWeek(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ThisMonth => ApplyThisMonth(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ThisWeek => ApplyThisWeek(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Today => ApplyToday(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Tomorrow => ApplyTomorrow(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Yesterday => ApplyYesterday(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Top => ApplyTop(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.TopPercent => ApplyTopPercent(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Bottom => ApplyBottom(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.BottomPercent => ApplyBottomPercent(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.AboveAverage => ApplyAboveAverage(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.AboveOrEqualAverage => ApplyAboveOrEqualAverage(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.BelowAverage => ApplyBelowAverage(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.BelowOrEqualAverage => ApplyBelowOrEqualAverage(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.AboveStdDev => ApplyAboveStdDev(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.BelowStdDev => ApplyBelowStdDev(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.Between => ApplyBetween(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NotBetween => ApplyNotBetween(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.NotContains => ApplyNotContains(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.TwoColorScale => ApplyTwoColorScale(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ThreeColorScale => ApplyThreeColorScale(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.DataBar => ApplyDataBar(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.ThreeIconSet => ApplyThreeIconSet(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.FourIconSet => ApplyFourIconSet(source, worksheet, newAddress),
                eExcelConditionalFormattingRuleType.FiveIconSet => ApplyFiveIconSet(source, worksheet, newAddress),
                _ => ApplyFallback(source, worksheet, newAddress)
            };
        }
        catch
        {
            return ApplyFallback(source, worksheet, newAddress);
        }
    }

    #region Extract

    private static void ExtractCellStyle(IExcelConditionalFormattingRule source, ConditionalFormattingRule target)
    {
        var style = source.Style;
        if (style == null)
        {
            return;
        }

        // Fill
        if (style.Fill != null)
        {
            target.FillStyle = new CfFillStyle
            {
                PatternType = style.Fill.PatternType ?? ExcelFillStyle.None,
                BackgroundColor = style.Fill.BackgroundColor?.Color ?? Color.Empty,
                ForegroundColor = style.Fill.PatternColor?.Color ?? Color.Empty
            };
        }

        // Font
        if (style.Font is ExcelDxfFont font)
        {
            target.FontStyle = new CfFontStyle
            {
                Bold = font.Bold ?? false,
                Italic = font.Italic ?? false,
                Underline = (font.Underline ?? ExcelUnderLineType.None) != ExcelUnderLineType.None,
                Color = font.Color?.Color ?? Color.Empty,
                Size = font.Size ?? 0f,
                Name = font.Name ?? string.Empty
            };
        }
        else if (style.Font != null)
        {
            // Fallback for ExcelDxfFontBase
            target.FontStyle = new CfFontStyle
            {
                Bold = style.Font.Bold ?? false,
                Italic = style.Font.Italic ?? false,
                Underline = (style.Font.Underline ?? ExcelUnderLineType.None) != ExcelUnderLineType.None,
                Color = style.Font.Color?.Color ?? Color.Empty,
                Size = 0f,
                Name = string.Empty
            };
        }

        // Border
        if (style.Border != null)
        {
            target.BorderStyle = new CfBorderStyle
            {
                Top = ExtractBorderEdge(style.Border.Top),
                Bottom = ExtractBorderEdge(style.Border.Bottom),
                Left = ExtractBorderEdge(style.Border.Left),
                Right = ExtractBorderEdge(style.Border.Right)
            };
        }
    }

    private static CfBorderEdge ExtractBorderEdge(ExcelDxfBorderItem edge)
    {
        return new CfBorderEdge
        {
            Style = edge?.Style ?? ExcelBorderStyle.None,
            Color = edge?.Color?.Color ?? Color.Empty
        };
    }

    private static void ExtractTwoColorScale(IExcelConditionalFormattingRule source, ConditionalFormattingRule target)
    {
        var scale = source.As.TwoColorScale;
        if (scale == null)
        {
            return;
        }

        target.ColorScaleStops =
        [
            new CfColorStop
            {
                Color = scale.LowValue?.Color ?? Color.Empty,
                Type = scale.LowValue?.Type ?? eExcelConditionalFormattingValueObjectType.Min,
                Value = scale.LowValue?.Value
            },
            new CfColorStop
            {
                Color = scale.HighValue?.Color ?? Color.Empty,
                Type = scale.HighValue?.Type ?? eExcelConditionalFormattingValueObjectType.Max,
                Value = scale.HighValue?.Value
            }
        ];
    }

    private static void ExtractThreeColorScale(IExcelConditionalFormattingRule source, ConditionalFormattingRule target)
    {
        var scale = source.As.ThreeColorScale;
        if (scale == null)
        {
            return;
        }

        target.ColorScaleStops =
        [
            new CfColorStop
            {
                Color = scale.LowValue?.Color ?? Color.Empty,
                Type = scale.LowValue?.Type ?? eExcelConditionalFormattingValueObjectType.Min,
                Value = scale.LowValue?.Value
            },
            new CfColorStop
            {
                Color = scale.MiddleValue?.Color ?? Color.Empty,
                Type = scale.MiddleValue?.Type ?? eExcelConditionalFormattingValueObjectType.Percent,
                Value = scale.MiddleValue?.Value
            },
            new CfColorStop
            {
                Color = scale.HighValue?.Color ?? Color.Empty,
                Type = scale.HighValue?.Type ?? eExcelConditionalFormattingValueObjectType.Max,
                Value = scale.HighValue?.Value
            }
        ];
    }

    private static void ExtractDataBar(IExcelConditionalFormattingRule source, ConditionalFormattingRule target)
    {
        var bar = source.As.DataBar;
        if (bar == null)
        {
            return;
        }

        target.DataBarSettings = new CfDataBarSettings
        {
            Color = bar.Color,
            MinValue = new CfValueBound
            {
                Type = bar.LowValue?.Type ?? eExcelConditionalFormattingValueObjectType.Min,
                Value = bar.LowValue?.Value
            },
            MaxValue = new CfValueBound
            {
                Type = bar.HighValue?.Type ?? eExcelConditionalFormattingValueObjectType.Max,
                Value = bar.HighValue?.Value
            },
            ShowValue = bar.ShowValue,
            AxisColor = bar.AxisColor?.Color,
            AxisPosition = (int?)bar.AxisPosition,
            BorderColor = bar.BorderColor?.Color,
            Direction = (int?)bar.Direction
        };
    }

    private static void ExtractIconSet(IExcelConditionalFormattingRule source, ConditionalFormattingRule target)
    {
        target.IconSetSettings = new CfIconSetSettings
        {
            ShowValue = true,
            Reverse = false,
            Criteria = []
        };

        switch (source.Type)
        {
            case eExcelConditionalFormattingRuleType.ThreeIconSet:
            {
                var iconSet = source.As.ThreeIconSet;
                if (iconSet == null)
                {
                    return;
                }

                target.IconSetSettings.ShowValue = iconSet.ShowValue;
                target.IconSetSettings.Reverse = iconSet.Reverse;
                target.IconSetSettings.IconSetType = (int)(object)iconSet.IconSet;

                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon1);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon2);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon3);

                break;
            }

            case eExcelConditionalFormattingRuleType.FourIconSet:
            {
                var iconSet = source.As.FourIconSet;
                if (iconSet == null)
                {
                    return;
                }

                target.IconSetSettings.ShowValue = iconSet.ShowValue;
                target.IconSetSettings.Reverse = iconSet.Reverse;
                target.IconSetSettings.IconSetType = (int)(object)iconSet.IconSet;

                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon1);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon2);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon3);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon4);

                break;
            }

            case eExcelConditionalFormattingRuleType.FiveIconSet:
            {
                var iconSet = source.As.FiveIconSet;
                if (iconSet == null)
                {
                    return;
                }

                target.IconSetSettings.ShowValue = iconSet.ShowValue;
                target.IconSetSettings.Reverse = iconSet.Reverse;
                target.IconSetSettings.IconSetType = (int)(object)iconSet.IconSet;

                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon1);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon2);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon3);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon4);
                ExtractIconCriterion(target.IconSetSettings.Criteria, iconSet.Icon5);

                break;
            }
        }
    }

    private static void ExtractIconCriterion(List<CfIconCriterion> criteria, ExcelConditionalFormattingIconDataBarValue value)
    {
        if (value == null)
        {
            return;
        }

        criteria.Add(new CfIconCriterion
        {
            Type = value.Type,
            Value = value.Value,
            Operator = value.GreaterThanOrEqualTo ? 0 : null
        });
    }

    #endregion

    #region Apply - Cell Style Rules

    private static IExcelConditionalFormattingRule ApplyCellStyle(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress, Func<IExcelConditionalFormattingRule> factory)
    {
        var cf = factory();
        ApplyCommonProperties(source, cf);
        ApplyStyle(source, cf);
        return cf;
    }

    private static IExcelConditionalFormattingRule ApplyExpression(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddExpression(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyGreaterThan(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddGreaterThan(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyGreaterThanOrEqual(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddGreaterThanOrEqual(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyLessThan(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddLessThan(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyLessThanOrEqual(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddLessThanOrEqual(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyEqual(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddEqual(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNotEqual(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNotEqual(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyContainsText(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddContainsText(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNotContainsText(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNotContainsText(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBeginsWith(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBeginsWith(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyEndsWith(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddEndsWith(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyContainsBlanks(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddContainsBlanks(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNotContainsBlanks(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNotContainsBlanks(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyContainsErrors(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddContainsErrors(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNotContainsErrors(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNotContainsErrors(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyDuplicateValues(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddDuplicateValues(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyUniqueValues(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddUniqueValues(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyLast7Days(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddLast7Days(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyLastMonth(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddLastMonth(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyLastWeek(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddLastWeek(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNextMonth(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNextMonth(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNextWeek(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNextWeek(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyThisMonth(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddThisMonth(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyThisWeek(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddThisWeek(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyToday(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddToday(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyTomorrow(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddTomorrow(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyYesterday(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddYesterday(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyTop(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddTop(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyTopPercent(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddTopPercent(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBottom(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBottom(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBottomPercent(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBottomPercent(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyAboveAverage(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddAboveAverage(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyAboveOrEqualAverage(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddAboveOrEqualAverage(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBelowAverage(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBelowAverage(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBelowOrEqualAverage(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBelowOrEqualAverage(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyAboveStdDev(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddAboveStdDev(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBelowStdDev(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBelowStdDev(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyBetween(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddBetween(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNotBetween(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddNotBetween(newAddress));
    }

    private static IExcelConditionalFormattingRule ApplyNotContains(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        return ApplyCellStyle(source, worksheet, newAddress, () => worksheet.ConditionalFormatting.AddTextContains(newAddress));
    }

    #endregion

    #region Apply - Color Scale

    private static IExcelConditionalFormattingRule ApplyTwoColorScale(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var cf = worksheet.ConditionalFormatting.AddTwoColorScale(newAddress);
        ApplyCommonProperties(source, cf);

        if (source.ColorScaleStops?.Count >= 2)
        {
            cf.LowValue.Color = source.ColorScaleStops[0].Color;
            cf.LowValue.Type = source.ColorScaleStops[0].Type;
            cf.LowValue.Value = source.ColorScaleStops[0].Value ?? 0;

            cf.HighValue.Color = source.ColorScaleStops[1].Color;
            cf.HighValue.Type = source.ColorScaleStops[1].Type;
            cf.HighValue.Value = source.ColorScaleStops[1].Value ?? 0;
        }

        return cf;
    }

    private static IExcelConditionalFormattingRule ApplyThreeColorScale(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var cf = worksheet.ConditionalFormatting.AddThreeColorScale(newAddress);
        ApplyCommonProperties(source, cf);

        if (source.ColorScaleStops?.Count >= 3)
        {
            cf.LowValue.Color = source.ColorScaleStops[0].Color;
            cf.LowValue.Type = source.ColorScaleStops[0].Type;
            cf.LowValue.Value = source.ColorScaleStops[0].Value ?? 0;

            cf.MiddleValue.Color = source.ColorScaleStops[1].Color;
            cf.MiddleValue.Type = source.ColorScaleStops[1].Type;
            cf.MiddleValue.Value = source.ColorScaleStops[1].Value ?? 0;

            cf.HighValue.Color = source.ColorScaleStops[2].Color;
            cf.HighValue.Type = source.ColorScaleStops[2].Type;
            cf.HighValue.Value = source.ColorScaleStops[2].Value ?? 0;
        }

        return cf;
    }

    #endregion

    #region Apply - DataBar

    private static IExcelConditionalFormattingRule ApplyDataBar(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var settings = source.DataBarSettings;
        var color = settings?.Color ?? Color.Blue;
        var cf = worksheet.ConditionalFormatting.AddDatabar(newAddress, color);
        ApplyCommonProperties(source, cf);

        if (settings != null)
        {
            cf.ShowValue = settings.ShowValue;

            if (settings.MinValue != null)
            {
                cf.LowValue.Type = settings.MinValue.Type;
                cf.LowValue.Value = settings.MinValue.Value ?? 0;
            }

            if (settings.MaxValue != null)
            {
                cf.HighValue.Type = settings.MaxValue.Type;
                cf.HighValue.Value = settings.MaxValue.Value ?? 0;
            }

            if (settings.AxisColor.HasValue)
            {
                cf.AxisColor.Color = settings.AxisColor.Value;
            }

            if (settings.AxisPosition.HasValue)
            {
                cf.AxisPosition = (eExcelDatabarAxisPosition)settings.AxisPosition.Value;
            }

            if (settings.BorderColor.HasValue)
            {
                cf.BorderColor.Color = settings.BorderColor.Value;
            }

            if (settings.Direction.HasValue)
            {
                cf.Direction = (eDatabarDirection)settings.Direction.Value;
            }
        }

        return cf;
    }

    #endregion

    #region Apply - IconSet

    private static IExcelConditionalFormattingRule ApplyThreeIconSet(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var settings = source.IconSetSettings;
        var iconSetType = settings != null
            ? (eExcelconditionalFormatting3IconsSetType)settings.IconSetType
            : eExcelconditionalFormatting3IconsSetType.TrafficLights1;
        var cf = worksheet.ConditionalFormatting.AddThreeIconSet(newAddress, iconSetType);
        ApplyCommonProperties(source, cf);
        ApplyIconSetSettings(cf, settings);
        return cf;
    }

    private static IExcelConditionalFormattingRule ApplyFourIconSet(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var settings = source.IconSetSettings;
        var iconSetType = settings != null
            ? (eExcelconditionalFormatting4IconsSetType)settings.IconSetType
            : eExcelconditionalFormatting4IconsSetType.RedToBlack;
        var cf = worksheet.ConditionalFormatting.AddFourIconSet(newAddress, iconSetType);
        ApplyCommonProperties(source, cf);
        ApplyIconSetSettings(cf, settings);
        return cf;
    }

    private static IExcelConditionalFormattingRule ApplyFiveIconSet(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var settings = source.IconSetSettings;
        var iconSetType = settings != null
            ? (eExcelconditionalFormatting5IconsSetType)settings.IconSetType
            : eExcelconditionalFormatting5IconsSetType.Quarters;
        var cf = worksheet.ConditionalFormatting.AddFiveIconSet(newAddress, iconSetType);
        ApplyCommonProperties(source, cf);
        ApplyIconSetSettings(cf, settings);
        return cf;
    }

    private static void ApplyIconSetSettings<T>(T cf, CfIconSetSettings settings) where T : IExcelConditionalFormattingRule
    {
        if (settings == null)
        {
            return;
        }

        // Use reflection to access ShowValue, Reverse, and icon properties
        var showValueProp = typeof(T).GetProperty("ShowValue");
        var reverseProp = typeof(T).GetProperty("Reverse");

        showValueProp?.SetValue(cf, settings.ShowValue);
        reverseProp?.SetValue(cf, settings.Reverse);

        if (settings.Criteria != null)
        {
            for (var i = 0; i < settings.Criteria.Count; i++)
            {
                var criterion = settings.Criteria[i];
                var iconProp = typeof(T).GetProperty("Icon" + (i + 1));
                if (iconProp == null)
                {
                    continue;
                }

                var iconValue = iconProp.GetValue(cf);
                if (iconValue == null)
                {
                    continue;
                }

                var typeProp = iconValue.GetType().GetProperty("Type");
                var valueProp = iconValue.GetType().GetProperty("Value");
                var greaterThanOrEqualToProp = iconValue.GetType().GetProperty("GreaterThanOrEqualTo");

                typeProp?.SetValue(iconValue, criterion.Type);
                valueProp?.SetValue(iconValue, criterion.Value ?? 0);

                if (greaterThanOrEqualToProp != null && criterion.Operator.HasValue)
                {
                    greaterThanOrEqualToProp.SetValue(iconValue, criterion.Operator.Value == 0);
                }
            }
        }
    }

    #endregion

    #region Common Helpers

    private static void ApplyCommonProperties(ConditionalFormattingRule source, IExcelConditionalFormattingRule target)
    {
        target.Priority = source.Priority;
        target.StopIfTrue = source.StopIfTrue;

        if (target is IExcelConditionalFormattingWithFormula cfFormula && !string.IsNullOrEmpty(source.Formula))
        {
            cfFormula.Formula = source.Formula;
        }

        if (target is IExcelConditionalFormattingWithFormula2 cfFormula2 && !string.IsNullOrEmpty(source.Formula2))
        {
            cfFormula2.Formula2 = source.Formula2;
        }
    }

    private static void ApplyStyle(ConditionalFormattingRule source, IExcelConditionalFormattingRule target)
    {
        if (target.Style == null)
        {
            return;
        }

        if (source.FillStyle != null)
        {
            target.Style.Fill.PatternType = source.FillStyle.PatternType;
            target.Style.Fill.BackgroundColor.Color = source.FillStyle.BackgroundColor;
            target.Style.Fill.PatternColor.Color = source.FillStyle.ForegroundColor;
        }

        if (source.FontStyle != null)
        {
            if (target.Style.Font is ExcelDxfFont font)
            {
                font.Bold = source.FontStyle.Bold;
                font.Italic = source.FontStyle.Italic;
                font.Underline = source.FontStyle.Underline ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
                font.Color.Color = source.FontStyle.Color;
                font.Size = source.FontStyle.Size;
                font.Name = source.FontStyle.Name;
            }
            else
            {
                target.Style.Font.Bold = source.FontStyle.Bold;
                target.Style.Font.Italic = source.FontStyle.Italic;
                target.Style.Font.Underline = source.FontStyle.Underline ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
                target.Style.Font.Color.Color = source.FontStyle.Color;
            }
        }

        if (source.BorderStyle != null)
        {
            ApplyBorderEdge(target.Style.Border.Top, source.BorderStyle.Top);
            ApplyBorderEdge(target.Style.Border.Bottom, source.BorderStyle.Bottom);
            ApplyBorderEdge(target.Style.Border.Left, source.BorderStyle.Left);
            ApplyBorderEdge(target.Style.Border.Right, source.BorderStyle.Right);
        }
    }

    private static void ApplyBorderEdge(ExcelDxfBorderItem target, CfBorderEdge source)
    {
        if (source == null || target == null)
        {
            return;
        }

        target.Style = source.Style;
        target.Color.Color = source.Color;
    }

    private static IExcelConditionalFormattingRule ApplyFallback(ConditionalFormattingRule source, ExcelWorksheet worksheet, string newAddress)
    {
        var cf = worksheet.ConditionalFormatting.AddExpression(newAddress);
        cf.Formula = source.Formula;
        cf.Style.Fill.BackgroundColor.Color = Color.Red;
        cf.Priority = source.Priority;
        cf.StopIfTrue = source.StopIfTrue;
        return cf;
    }

    #endregion
}
