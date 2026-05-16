using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using EPPlus.Report.Model;
using OfficeOpenXml;

namespace EPPlus.Report.Parsing;

/// <summary>
///     Scans an Excel worksheet for named ranges that can be interpreted as loop nodes.
/// </summary>
public class NamedRangeScanner
{
    private static readonly Regex ServiceTagRegex = new(@"<<(\w+)>>", RegexOptions.Compiled);
    private static readonly Regex GroupTagRegex = new(@"<<group\s+(\w+)(?:\s+(asc|desc))?>>", RegexOptions.Compiled);
    private static readonly Regex ExpressionRegex = new(@"\{\{(.+?)\}\}", RegexOptions.Compiled);

    /// <summary>
    ///     Scans the worksheet for named ranges and converts them into <see cref="NamedRangeLoopNode" /> instances.
    /// </summary>
    /// <param name="worksheet">The Excel worksheet to scan.</param>
    /// <param name="errors">A collection to populate with scanning errors.</param>
    /// <returns>A list of <see cref="NamedRangeLoopNode" /> instances found in the worksheet.</returns>
    public List<NamedRangeLoopNode> Scan(ExcelWorksheet worksheet, TemplateErrors errors)
    {
        var nodes = new List<NamedRangeLoopNode>();

        if (worksheet?.Names == null)
        {
            return nodes;
        }

        // Scan worksheet-level named ranges
        foreach (var namedRange in worksheet.Names)
        {
            var node = TryConvertToLoopNode(worksheet, namedRange, errors);
            if (node != null)
            {
                nodes.Add(node);
            }
        }

        // Scan workbook-level named ranges that target this worksheet
        if (worksheet.Workbook?.Names != null)
        {
            foreach (var namedRange in worksheet.Workbook.Names)
            {
                if (namedRange.Worksheet != null && namedRange.Worksheet.Name == worksheet.Name)
                {
                    var node = TryConvertToLoopNode(worksheet, namedRange, errors);
                    if (node != null)
                    {
                        nodes.Add(node);
                    }
                }
            }
        }

        return nodes;
    }

    private NamedRangeLoopNode TryConvertToLoopNode(ExcelWorksheet worksheet, ExcelNamedRange namedRange,
        TemplateErrors errors)
    {
        if (namedRange.Start == null || namedRange.End == null)
        {
            return null;
        }

        var startRow = namedRange.Start.Row;
        var startCol = namedRange.Start.Column;
        var endRow = namedRange.End.Row;
        var endCol = namedRange.End.Column;

        var rowCount = endRow - startRow + 1;
        var colCount = endCol - startCol + 1;

        // Vertical table: at least 2 rows
        if (rowCount < 2)
        {
            return null;
        }

        var loopNode = new NamedRangeLoopNode
        {
            RangeName = namedRange.Name,
            CollectionName = namedRange.Name,
            Row = startRow,
            Column = startCol,
            EndRow = endRow,
            EndColumn = endCol,
            IsHorizontal = false,
            RawContent = $"<<namedrange {namedRange.Name}>>"
        };

        // Parse service row tags first to determine ServiceRowCount
        var serviceRow = endRow;
        var serviceTags = new List<ServiceTag>();
        var groupByDefinitions = new List<GroupByDefinition>();

        for (var c = startCol; c <= endCol; c++)
        {
            var cellValue = worksheet.Cells[serviceRow, c].Text;
            if (string.IsNullOrWhiteSpace(cellValue))
            {
                continue;
            }

            var groupMatch = GroupTagRegex.Match(cellValue);
            if (groupMatch.Success)
            {
                groupByDefinitions.Add(new GroupByDefinition
                {
                    PropertyPath = groupMatch.Groups[1].Value,
                    Column = c,
                    Descending = groupMatch.Groups[2].Success &&
                                 groupMatch.Groups[2].Value.Equals("desc", StringComparison.OrdinalIgnoreCase)
                });
                continue;
            }

            var tagMatch = ServiceTagRegex.Match(cellValue);
            if (tagMatch.Success && !tagMatch.Groups[1].Value.Equals("group", StringComparison.OrdinalIgnoreCase))
            {
                serviceTags.Add(new ServiceTag
                {
                    TagName = tagMatch.Groups[1].Value.ToLowerInvariant(),
                    Row = serviceRow,
                    Column = c
                });
            }
        }

        // If no service row inside the named range, check the adjacent row below
        if (serviceTags.Count == 0 && groupByDefinitions.Count == 0)
        {
            var adjacentRow = endRow + 1;
            if (adjacentRow <= worksheet.Dimension?.End.Row)
            {
                var adjacentTags = new List<ServiceTag>();
                var adjacentGroups = new List<GroupByDefinition>();
                var hasNonServiceContent = false;

                for (var c = startCol; c <= endCol; c++)
                {
                    var cellValue = worksheet.Cells[adjacentRow, c].Text;
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        continue;
                    }

                    // Any expression means this is not a service row
                    if (ExpressionRegex.IsMatch(cellValue))
                    {
                        hasNonServiceContent = true;
                        break;
                    }

                    var groupMatch = GroupTagRegex.Match(cellValue);
                    if (groupMatch.Success)
                    {
                        adjacentGroups.Add(new GroupByDefinition
                        {
                            PropertyPath = groupMatch.Groups[1].Value,
                            Column = c,
                            Descending = groupMatch.Groups[2].Success &&
                                         groupMatch.Groups[2].Value.Equals("desc", StringComparison.OrdinalIgnoreCase)
                        });
                        continue;
                    }

                    var tagMatch = ServiceTagRegex.Match(cellValue);
                    if (tagMatch.Success && !tagMatch.Groups[1].Value.Equals("group", StringComparison.OrdinalIgnoreCase))
                    {
                        adjacentTags.Add(new ServiceTag
                        {
                            TagName = tagMatch.Groups[1].Value.ToLowerInvariant(),
                            Row = adjacentRow,
                            Column = c
                        });
                    }
                    else
                    {
                        hasNonServiceContent = true;
                        break;
                    }
                }

                if (!hasNonServiceContent && (adjacentTags.Count > 0 || adjacentGroups.Count > 0))
                {
                    endRow = adjacentRow;
                    loopNode.EndRow = endRow;
                    serviceTags = adjacentTags;
                    groupByDefinitions = adjacentGroups;
                }
            }
        }

        loopNode.ServiceTags = serviceTags;
        loopNode.GroupByDefinitions = groupByDefinitions;
        loopNode.ServiceRowCount = serviceTags.Count > 0 || groupByDefinitions.Count > 0 ? 1 : 0;

        // Parse data rows (all except service rows)
        var dataEndRow = endRow - loopNode.ServiceRowCount;
        var children = new List<TemplateNode>();

        for (var r = startRow; r <= dataEndRow; r++)
        {
            for (var c = startCol; c <= endCol; c++)
            {
                var cellValue = worksheet.Cells[r, c].Text;
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    continue;
                }

                var node = CreateNode(r, c, cellValue);
                if (node != null)
                {
                    children.Add(node);
                }
            }
        }

        // Compute header row count (leading rows without expressions)
        var headerRowCount = 0;
        for (var r = startRow; r <= dataEndRow; r++)
        {
            var hasExpression = false;
            for (var c = startCol; c <= endCol; c++)
            {
                var cellValue = worksheet.Cells[r, c].Text;
                if (ExpressionRegex.IsMatch(cellValue))
                {
                    hasExpression = true;
                    break;
                }
            }

            if (!hasExpression)
            {
                headerRowCount++;
            }
            else
            {
                break;
            }
        }

        loopNode.HeaderRowCount = headerRowCount;
        loopNode.Children = children;

        return loopNode;
    }

    private TemplateNode CreateNode(int row, int col, string cellValue)
    {
        var exprMatch = ExpressionRegex.Match(cellValue);
        if (exprMatch.Success)
        {
            return new ExpressionNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                ExpressionPath = exprMatch.Groups[1].Value.Trim()
            };
        }

        return new TextNode
        {
            Row = row,
            Column = col,
            RawContent = cellValue
        };
    }
}