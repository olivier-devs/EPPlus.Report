using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using EPPlus.Report.Model;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace EPPlus.Report.Parsing;

/// <summary>
///     Parses an Excel worksheet into a <see cref="Template" /> AST by scanning cells for directives such as loops,
///     conditions, groups, and expressions.
/// </summary>
public class TemplateParser : ITemplateParser
{
    private static readonly Regex ExpressionRegex = new(@"\{\{(.+?)\}\}", RegexOptions.Compiled);
    private static readonly Regex LoopStartRegex = new(@"<<foreach\s+(\w+)>>", RegexOptions.Compiled);
    private static readonly Regex LoopEndRegex = new(@"<</foreach>>", RegexOptions.Compiled);
    private static readonly Regex IfStartRegex = new(@"<<if\s+(\w+)>>", RegexOptions.Compiled);
    private static readonly Regex IfEndRegex = new(@"<</if>>", RegexOptions.Compiled);
    private static readonly Regex GroupStartRegex = new(@"<<group\s+(\w+)\s+by\s+(.+?)>>", RegexOptions.Compiled);
    private static readonly Regex GroupEndRegex = new(@"<</group>>", RegexOptions.Compiled);
    private static readonly Regex SumRegex = new(@"<<sum\s+(\w+)>>", RegexOptions.Compiled);
    private static readonly Regex CountRegex = new(@"<<count\s+(\w+)>>", RegexOptions.Compiled);

    private static readonly Regex FunctionRegex =
        new(@"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(\s*([A-Za-z_][A-Za-z0-9_\.]*)\s*\)\s*\}\}", RegexOptions.Compiled);

    /// <summary>
    ///     Parses the specified worksheet and builds a <see cref="Template" /> representing the template structure.
    /// </summary>
    /// <param name="worksheet">The Excel worksheet to parse.</param>
    /// <param name="errors">A collection to populate with parsing errors.</param>
    /// <returns>The parsed <see cref="Template" />.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="worksheet" /> is null.</exception>
    public Template Parse(ExcelWorksheet worksheet, TemplateErrors errors)
    {
        if (worksheet is null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        errors ??= [];

        var template = new Template();

        if (worksheet.Dimension == null)
        {
            return template;
        }

        for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
        {
            for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var cellValue = worksheet.Cells[row, col].Text;
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    continue;
                }

                var node = ParseNode(worksheet, ref row, col, cellValue, errors);
                if (node != null)
                {
                    template.Nodes.Add(node);
                }
            }
        }

        // Scan for named range loops
        var scanner = new NamedRangeScanner();
        var namedRangeLoops = scanner.Scan(worksheet, errors);

        foreach (var nrLoop in namedRangeLoops)
        {
            // Check for overlap with existing block nodes (classic foreach/if takes priority)
            var overlaps = false;
            foreach (var existingNode in template.Nodes)
            {
                if (existingNode is LoopNode existingLoop)
                {
                    var rowOverlap = nrLoop.Row <= existingLoop.EndRow && nrLoop.EndRow >= existingLoop.Row;
                    var colOverlap = nrLoop.Column <= existingLoop.Column && nrLoop.EndColumn >= existingLoop.Column;
                    var sameCollection = string.Equals(nrLoop.CollectionName, existingLoop.CollectionName,
                        StringComparison.OrdinalIgnoreCase);
                    if (rowOverlap && (colOverlap || sameCollection))
                    {
                        overlaps = true;
                        break;
                    }
                }
                else if (existingNode is IfNode existingIf)
                {
                    var rowOverlap = nrLoop.Row <= existingIf.EndRow && nrLoop.EndRow >= existingIf.Row;
                    var colOverlap = nrLoop.Column <= existingIf.Column && nrLoop.EndColumn >= existingIf.Column;
                    if (rowOverlap && colOverlap)
                    {
                        overlaps = true;
                        break;
                    }
                }
            }

            if (!overlaps)
            {
                // Find existing nodes inside the named range that the scanner may have missed
                var nodesInsideRange = template.Nodes
                    .Where(n => n.Row >= nrLoop.Row && n.Row <= nrLoop.EndRow &&
                                n.Column >= nrLoop.Column && n.Column <= nrLoop.EndColumn)
                    .ToList();

                foreach (var existingNode in nodesInsideRange)
                {
                    var scannerChildIndex = nrLoop.Children.FindIndex(c =>
                        c.Row == existingNode.Row && c.Column == existingNode.Column);
                    if (scannerChildIndex >= 0)
                    {
                        // The scanner created a generic TextNode for a directive it didn't recognise.
                        // Replace it with the properly typed node from the first parsing pass.
                        if (nrLoop.Children[scannerChildIndex] is TextNode && existingNode is not TextNode)
                        {
                            nrLoop.Children[scannerChildIndex] = existingNode;
                        }
                    }
                    else
                    {
                        nrLoop.Children.Add(existingNode);
                    }
                }

                // Remove any existing nodes that fall inside the named range
                template.Nodes.RemoveAll(n =>
                    n.Row >= nrLoop.Row && n.Row <= nrLoop.EndRow &&
                    n.Column >= nrLoop.Column && n.Column <= nrLoop.EndColumn);
                template.Nodes.Add(nrLoop);
            }
        }

        AssociateConditionalFormatting(worksheet, template);

        return template;
    }

    private TemplateNode ParseNode(ExcelWorksheet worksheet, ref int row, int col, string cellValue,
        TemplateErrors errors)
    {
        var groupStartMatch = GroupStartRegex.Match(cellValue);
        if (groupStartMatch.Success)
        {
            return ParseBlock(worksheet, ref row, col, cellValue, groupStartMatch.Groups[1].Value, BlockType.Group,
                errors, groupStartMatch.Groups[2].Value);
        }

        var loopStartMatch = LoopStartRegex.Match(cellValue);
        if (loopStartMatch.Success)
        {
            return ParseBlock(worksheet, ref row, col, cellValue, loopStartMatch.Groups[1].Value, BlockType.Loop,
                errors);
        }

        var ifStartMatch = IfStartRegex.Match(cellValue);
        if (ifStartMatch.Success)
        {
            return ParseBlock(worksheet, ref row, col, cellValue, ifStartMatch.Groups[1].Value, BlockType.If, errors);
        }

        var sumMatch = SumRegex.Match(cellValue);
        if (sumMatch.Success)
        {
            return new AggregationNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                AggregationType = "sum",
                PropertyName = sumMatch.Groups[1].Value
            };
        }

        var countMatch = CountRegex.Match(cellValue);
        if (countMatch.Success)
        {
            return new AggregationNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                AggregationType = "count",
                PropertyName = countMatch.Groups[1].Value
            };
        }

        var funcMatch = FunctionRegex.Match(cellValue);
        if (funcMatch.Success)
        {
            return new ExpressionNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                FunctionName = funcMatch.Groups[1].Value.Trim(),
                ExpressionPath = funcMatch.Groups[2].Value.Trim()
            };
        }

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

    private TemplateNode ParseBlock(ExcelWorksheet worksheet, ref int row, int col, string cellValue, string expression,
        BlockType blockType, TemplateErrors errors, string groupByExpression = null)
    {
        TemplateNode blockNode = blockType switch
        {
            BlockType.Loop => new LoopNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                CollectionName = expression
            },
            BlockType.If => new IfNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                ConditionExpression = expression
            },
            BlockType.Group => new GroupNode
            {
                Row = row,
                Column = col,
                RawContent = cellValue,
                CollectionName = expression,
                GroupByPaths = ParseGroupByPaths(groupByExpression),
                Options = new GroupOptions
                    { Descending = groupByExpression?.EndsWith(" desc", StringComparison.OrdinalIgnoreCase) == true }
            },
            _ => throw new ArgumentException($"Unknown block type: {blockType}")
        };

        var children = new List<TemplateNode>();
        row++;

        while (row <= worksheet.Dimension.End.Row)
        {
            var nextCellValue = worksheet.Cells[row, col].Text;

            if (LoopEndRegex.IsMatch(nextCellValue) || IfEndRegex.IsMatch(nextCellValue) ||
                GroupEndRegex.IsMatch(nextCellValue))
            {
                row++;
                break;
            }

            if (string.IsNullOrWhiteSpace(nextCellValue))
            {
                row++;
                continue;
            }

            var child = ParseNode(worksheet, ref row, col, nextCellValue, errors);
            if (child != null)
            {
                children.Add(child);
            }

            row++;
        }

        if (row > worksheet.Dimension.End.Row)
        {
            errors.Add(new TemplateError
            {
                Message =
                    $"{(blockType == BlockType.Loop ? "foreach" : blockType == BlockType.If ? "if" : "group")} block '{expression}' is not closed",
                CellAddress = worksheet.Cells[blockNode.Row, col].Address,
                WorksheetName = worksheet.Name,
                Row = blockNode.Row,
                Column = col,
                Expression = cellValue,
                Type = ErrorType.Parsing
            });
        }

        row--; // Adjust for outer loop increment

        if (blockNode is GroupNode groupNode)
        {
            // Auto-detect subtotal template: last row with only aggregation nodes
            if (children.Count > 0)
            {
                var lastRow = children.Max(c => c.Row);
                var lastRowNodes = children.Where(c => c.Row == lastRow).ToList();
                if (lastRowNodes.All(n => n is AggregationNode))
                {
                    groupNode.SubtotalTemplate = lastRowNodes;
                    children = children.Where(c => c.Row != lastRow).ToList();
                }
            }

            groupNode.Children = children;
            groupNode.EndRow = row;
        }
        else if (blockNode is IfNode ifNode)
        {
            ifNode.Children = children;
            ifNode.EndRow = row;
        }
        else if (blockNode is LoopNode loopNode)
        {
            loopNode.Children = children;
            loopNode.EndRow = row;
        }

        return blockNode;
    }

    private static List<string> ParseGroupByPaths(string groupByExpression)
    {
        var parts = groupByExpression.Split(',');
        var paths = new List<string>();
        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (trimmed.EndsWith(" asc", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 4).Trim();
            }
            else if (trimmed.EndsWith(" desc", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 5).Trim();
            }

            if (!string.IsNullOrEmpty(trimmed))
            {
                paths.Add(trimmed);
            }
        }

        return paths;
    }

    private static void AssociateConditionalFormatting(ExcelWorksheet worksheet, Template template)
    {
        foreach (var node in template.Nodes)
        {
            var startRow = node.Row;
            var endRow = node.Row;

            if (node is GroupNode group)
            {
                endRow = group.EndRow;
            }
            else if (node is IfNode ifNode)
            {
                endRow = ifNode.EndRow;
            }
            else if (node is LoopNode loop)
            {
                endRow = loop.EndRow;
            }
            else
            {
                continue;
            }

            foreach (var cf in worksheet.ConditionalFormatting)
            {
                var cfAddress = cf.Address;
                if (cfAddress.Start.Row <= endRow && cfAddress.End.Row >= startRow)
                {
                    var rule = new ConditionalFormattingRule
                    {
                        Address = cf.Address.Address,
                        Type = cf.Type,
                        Priority = cf.Priority,
                        StopIfTrue = cf.StopIfTrue
                    };

                    if (cf is IExcelConditionalFormattingWithFormula cfFormula)
                    {
                        rule.Formula = cfFormula.Formula ?? string.Empty;
                    }

                    if (cf is IExcelConditionalFormattingWithFormula2 cfFormula2)
                    {
                        rule.Formula2 = cfFormula2.Formula2 ?? string.Empty;
                    }

                    switch (node)
                    {
                        case GroupNode gn: gn.ConditionalFormattingRules.Add(rule); break;
                        case IfNode @in: @in.ConditionalFormattingRules.Add(rule); break;
                        case LoopNode ln: ln.ConditionalFormattingRules.Add(rule); break;
                    }
                }
            }
        }
    }

    private enum BlockType
    {
        Loop,
        If,
        Group
    }
}