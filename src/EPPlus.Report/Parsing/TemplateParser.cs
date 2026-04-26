using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using EPPlus.Report.Model;
using OfficeOpenXml;

namespace EPPlus.Report.Parsing
{
    public class TemplateParser : ITemplateParser
    {
        private static readonly Regex _expressionRegex = new Regex(@"\{\{(.+?)\}\}", RegexOptions.Compiled);
        private static readonly Regex _loopStartRegex = new Regex(@"<<foreach\s+(\w+)>>", RegexOptions.Compiled);
        private static readonly Regex _loopEndRegex = new Regex(@"<</foreach>>", RegexOptions.Compiled);
        private static readonly Regex _ifStartRegex = new Regex(@"<<if\s+(\w+)>>", RegexOptions.Compiled);
        private static readonly Regex _ifEndRegex = new Regex(@"<</if>>", RegexOptions.Compiled);
        private static readonly Regex _groupStartRegex = new Regex(@"<<group\s+(\w+)\s+by\s+(.+?)>>", RegexOptions.Compiled);
        private static readonly Regex _groupEndRegex = new Regex(@"<</group>>", RegexOptions.Compiled);
        private static readonly Regex _sumRegex = new Regex(@"<<sum\s+(\w+)>>", RegexOptions.Compiled);
        private static readonly Regex _countRegex = new Regex(@"<<count\s+(\w+)>>", RegexOptions.Compiled);
        private static readonly Regex _functionRegex = new Regex(@"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(\s*([A-Za-z_][A-Za-z0-9_\.]*)\s*\)\s*\}\}", RegexOptions.Compiled);

        public Template Parse(ExcelWorksheet worksheet, TemplateErrors errors)
        {
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet));

            if (errors == null)
                errors = new TemplateErrors();

            var template = new Template();
            
            if (worksheet.Dimension == null)
                return template;

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text;
                    if (string.IsNullOrWhiteSpace(cellValue))
                        continue;

                    var node = ParseNode(worksheet, ref row, col, cellValue, errors);
                    if (node != null)
                        template.Nodes.Add(node);
                }
            }

            // Scan for named range loops
            var scanner = new NamedRangeScanner();
            var namedRangeLoops = scanner.Scan(worksheet, errors);

            foreach (var nrLoop in namedRangeLoops)
            {
                // Check for overlap with existing block nodes (classic foreach/if takes priority)
                bool overlaps = false;
                foreach (var existingNode in template.Nodes)
                {
                    if (existingNode is LoopNode existingLoop)
                    {
                        bool rowOverlap = nrLoop.Row <= existingLoop.EndRow && nrLoop.EndRow >= existingLoop.Row;
                        bool colOverlap = nrLoop.Column <= existingLoop.Column && nrLoop.EndColumn >= existingLoop.Column;
                        bool sameCollection = string.Equals(nrLoop.CollectionName, existingLoop.CollectionName, StringComparison.OrdinalIgnoreCase);
                        if (rowOverlap && (colOverlap || sameCollection))
                        {
                            overlaps = true;
                            break;
                        }
                    }
                    else if (existingNode is IfNode existingIf)
                    {
                        bool rowOverlap = nrLoop.Row <= existingIf.EndRow && nrLoop.EndRow >= existingIf.Row;
                        bool colOverlap = nrLoop.Column <= existingIf.Column && nrLoop.EndColumn >= existingIf.Column;
                        if (rowOverlap && colOverlap)
                        {
                            overlaps = true;
                            break;
                        }
                    }
                }

                if (!overlaps)
                {
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

        private TemplateNode ParseNode(ExcelWorksheet worksheet, ref int row, int col, string cellValue, TemplateErrors errors)
        {
            var groupStartMatch = _groupStartRegex.Match(cellValue);
            if (groupStartMatch.Success)
            {
                return ParseBlock(worksheet, ref row, col, cellValue, groupStartMatch.Groups[1].Value, BlockType.Group, errors, groupStartMatch.Groups[2].Value);
            }

            var loopStartMatch = _loopStartRegex.Match(cellValue);
            if (loopStartMatch.Success)
            {
                return ParseBlock(worksheet, ref row, col, cellValue, loopStartMatch.Groups[1].Value, BlockType.Loop, errors);
            }

            var ifStartMatch = _ifStartRegex.Match(cellValue);
            if (ifStartMatch.Success)
            {
                return ParseBlock(worksheet, ref row, col, cellValue, ifStartMatch.Groups[1].Value, BlockType.If, errors);
            }

            var sumMatch = _sumRegex.Match(cellValue);
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

            var countMatch = _countRegex.Match(cellValue);
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

            var funcMatch = _functionRegex.Match(cellValue);
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

            var exprMatch = _expressionRegex.Match(cellValue);
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

        private TemplateNode ParseBlock(ExcelWorksheet worksheet, ref int row, int col, string cellValue, string expression, BlockType blockType, TemplateErrors errors, string groupByExpression = null)
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
                    Options = new GroupOptions { Descending = groupByExpression?.EndsWith(" desc", StringComparison.OrdinalIgnoreCase) == true }
                },
                _ => throw new ArgumentException($"Unknown block type: {blockType}")
            };

            var children = new List<TemplateNode>();
            row++;

            while (row <= worksheet.Dimension.End.Row)
            {
                var nextCellValue = worksheet.Cells[row, col].Text;

                if (_loopEndRegex.IsMatch(nextCellValue) || _ifEndRegex.IsMatch(nextCellValue) || _groupEndRegex.IsMatch(nextCellValue))
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
                    children.Add(child);

                row++;
            }

            if (row > worksheet.Dimension.End.Row)
            {
                errors.Add(new TemplateError
                {
                    Message = $"{(blockType == BlockType.Loop ? "foreach" : blockType == BlockType.If ? "if" : "group")} block '{expression}' is not closed",
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

        private List<string> ParseGroupByPaths(string groupByExpression)
        {
            var parts = groupByExpression.Split(',');
            var paths = new List<string>();
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.EndsWith(" asc", StringComparison.OrdinalIgnoreCase))
                    trimmed = trimmed.Substring(0, trimmed.Length - 4).Trim();
                else if (trimmed.EndsWith(" desc", StringComparison.OrdinalIgnoreCase))
                    trimmed = trimmed.Substring(0, trimmed.Length - 5).Trim();
                if (!string.IsNullOrEmpty(trimmed))
                    paths.Add(trimmed);
            }
            return paths;
        }

        private void AssociateConditionalFormatting(ExcelWorksheet worksheet, Template template)
        {
            foreach (var node in template.Nodes)
            {
                int startRow = node.Row;
                int endRow = node.Row;

                if (node is GroupNode group) endRow = group.EndRow;
                else if (node is IfNode ifNode) endRow = ifNode.EndRow;
                else if (node is LoopNode loop) endRow = loop.EndRow;
                else continue;

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

                        if (cf is OfficeOpenXml.ConditionalFormatting.Contracts.IExcelConditionalFormattingWithFormula cfFormula)
                        {
                            rule.Formula = cfFormula.Formula?.ToString() ?? string.Empty;
                        }
                        if (cf is OfficeOpenXml.ConditionalFormatting.Contracts.IExcelConditionalFormattingWithFormula2 cfFormula2)
                        {
                            rule.Formula2 = cfFormula2.Formula2?.ToString() ?? string.Empty;
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
}
