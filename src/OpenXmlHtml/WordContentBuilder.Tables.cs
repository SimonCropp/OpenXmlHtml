static partial class WordContentBuilder
{
    static void BuildTable(IElement tableElement, FormatState format, List<OpenXmlElement> elements, WordBuildContext ctx)
    {
        IElement? caption = null;
        foreach (var child in tableElement.Children)
        {
            if (child.LocalName == "caption")
            {
                caption = child;
                break;
            }
        }

        if (caption != null)
        {
            var captionFormat = format;
            HtmlSegmentParser.ApplyElementFormatting(caption, "caption", ref captionFormat);
            ProcessChildren(caption, captionFormat, elements, ctx, false);
            FlushParagraph(elements, ctx);
        }

        var rows = GetTableRows(tableElement);
        if (rows.Count == 0)
        {
            return;
        }

        var columnCount = GetColumnCount(rows);
        var table = new Table();

        var defaultBorder = new BorderInfo(4, BorderValues.Single, "auto");

        var borderAttr = tableElement.GetAttribute("border");
        if (borderAttr == "0")
        {
            defaultBorder = new(0, BorderValues.None, null);
        }

        var tableStyle = tableElement.GetAttribute("style");
        Dictionary<string, string>? declarations = null;
        if (tableStyle != null)
        {
            declarations = StyleParser.Parse(tableStyle);

            if (declarations.TryGetValue("border", out var tableBorderCss))
            {
                defaultBorder = StyleParser.ParseBorder(tableBorderCss) ?? defaultBorder;
            }
        }

        var tableBorders = BorderEmitter.BuildTableBorders(defaultBorder);

        var tblPr = new TableProperties(
            new TableWidth
            {
                Width = "0",
                Type = TableWidthUnitValues.Auto
            },
            tableBorders);

        if (declarations != null)
        {
            if (declarations.TryGetValue("width", out var tableWidth))
            {
                var twips = StyleParser.ParseLengthToTwips(tableWidth);
                if (twips != null)
                {
                    tblPr.TableWidth = new()
                    {
                        Width = twips.Value.ToString(),
                        Type = TableWidthUnitValues.Dxa
                    };
                }
            }

            if (declarations.TryGetValue("background-color", out var tableBg) ||
                declarations.TryGetValue("background", out tableBg))
            {
                var parsed = ColorParser.Parse(tableBg);
                if (parsed != null)
                {
                    tblPr.Append(new Shading
                    {
                        Val = ShadingPatternValues.Clear,
                        Fill = parsed
                    });
                }
            }

            ApplyTableCellPadding(declarations, tblPr);
        }

        var cellPaddingAttr = tableElement.GetAttribute("cellpadding");
        if (cellPaddingAttr != null)
        {
            var twips = StyleParser.ParseLengthToTwips(cellPaddingAttr);
            if (twips != null)
            {
                tblPr.Append(PaddingHelper.BuildMargin<TableCellMarginDefault>(twips, twips, twips, twips));
            }
        }

        table.Append(tblPr);

        var columnWidths = GetColumnWidths(tableElement, columnCount);
        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            var gridCol = new GridColumn();
            if (columnWidths[i] is { } w)
            {
                gridCol.Width = w.ToString();
            }

            grid.Append(gridCol);
        }

        table.Append(grid);

        // Track rowspans: starting column index -> (remaining rows, colspan)
        var rowspanTracker = new Dictionary<int, (int Remaining, int Colspan)>();

        foreach (var row in rows)
        {
            var tr = new TableRow();

            var rowHeight = row.GetAttribute("style") is { } rowStyle
                ? StyleParser.ParseLengthToTwips(StyleParser.Parse(rowStyle).GetValueOrDefault("height") ?? "")
                : null;
            rowHeight ??= row.GetAttribute("height") is { } rh ? StyleParser.ParseLengthToTwips(rh) : null;
            if (rowHeight != null)
            {
                tr.Append(new TableRowProperties(
                    new TableRowHeight
                    {
                        Val = (uint) rowHeight.Value,
                        HeightType = HeightRuleValues.AtLeast
                    }));
            }

            var cells = GetCells(row);
            var cellIndex = 0;
            var colIndex = 0;

            while (colIndex < columnCount)
            {
                if (rowspanTracker.TryGetValue(colIndex, out var spanInfo))
                {
                    var contTcPr = new TableCellProperties(new VerticalMerge());
                    if (spanInfo.Colspan > 1)
                    {
                        contTcPr.Append(new GridSpan
                        {
                            Val = spanInfo.Colspan
                        });
                    }

                    tr.Append(new TableCell(contTcPr, new Paragraph()));

                    if (spanInfo.Remaining <= 1)
                    {
                        rowspanTracker.Remove(colIndex);
                    }
                    else
                    {
                        rowspanTracker[colIndex] = (spanInfo.Remaining - 1, spanInfo.Colspan);
                    }

                    colIndex += spanInfo.Colspan;
                    continue;
                }

                if (cellIndex >= cells.Count)
                {
                    tr.Append(new TableCell(new Paragraph()));
                    colIndex++;
                    continue;
                }

                var cellElement = cells[cellIndex];
                cellIndex++;

                var colspan = ParseIntAttribute(cellElement, "colspan", 1);
                var rowspan = ParseIntAttribute(cellElement, "rowspan", 1);

                int? cellColWidth = null;
                for (var i = 0; i < colspan && colIndex + i < columnWidths.Count; i++)
                {
                    if (columnWidths[colIndex + i] is { } cw)
                    {
                        cellColWidth = (cellColWidth ?? 0) + cw;
                    }
                }

                tr.Append(BuildTableCell(cellElement, format, ctx, colspan, rowspan > 1, cellColWidth));

                if (rowspan > 1)
                {
                    rowspanTracker[colIndex] = (rowspan - 1, colspan);
                }

                colIndex += colspan;
            }

            table.Append(tr);
        }

        elements.Add(table);
    }

    static TableCell BuildTableCell(IElement cellElement, FormatState format, WordBuildContext parentCtx, int colspan, bool isRowspanStart, int? colWidth)
    {
        var tc = new TableCell();
        TableCellProperties? tcPr = null;

        if (colspan > 1 || isRowspanStart)
        {
            tcPr = new();
            if (colspan > 1)
            {
                tcPr.Append(new GridSpan
                {
                    Val = colspan
                });
            }

            if (isRowspanStart)
            {
                tcPr.Append(new VerticalMerge
                {
                    Val = MergedCellValues.Restart
                });
            }
        }

        var cellStyle = cellElement.GetAttribute("style");
        if (cellStyle != null)
        {
            var declarations = StyleParser.Parse(cellStyle);
            tcPr = ApplyCellStyles(declarations, tcPr);
        }

        var bgColorAttr = cellElement.GetAttribute("bgcolor");
        if (bgColorAttr != null)
        {
            var parsed = ColorParser.Parse(bgColorAttr);
            if (parsed != null)
            {
                tcPr ??= new();
                tcPr.Append(new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Fill = parsed
                });
            }
        }

        var widthAttr = cellElement.GetAttribute("width");
        if (widthAttr != null)
        {
            var twips = StyleParser.ParseLengthToTwips(widthAttr);
            if (twips != null)
            {
                tcPr ??= new();
                tcPr.Append(new TableCellWidth
                {
                    Width = twips.Value.ToString(),
                    Type = TableWidthUnitValues.Dxa
                });
            }
        }

        if (colWidth != null && tcPr?.GetFirstChild<TableCellWidth>() == null)
        {
            tcPr ??= new();
            tcPr.Append(new TableCellWidth
            {
                Width = colWidth.Value.ToString(),
                Type = TableWidthUnitValues.Dxa
            });
        }

        if (tcPr != null)
        {
            tc.Append(tcPr);
        }

        var cellFormat = format;
        HtmlSegmentParser.ApplyElementFormatting(cellElement, cellElement.LocalName, ref cellFormat);

        var cellElements = new List<OpenXmlElement>();
        var cellCtx = new WordBuildContext
        {
            MainPart = parentCtx.MainPart,
            ImageIndex = parentCtx.ImageIndex,
            Settings = parentCtx.Settings,
            StyleMap = parentCtx.StyleMap,
            BulletAbstractNumId = parentCtx.BulletAbstractNumId,
            NextNumId = parentCtx.NextNumId
        };
        ProcessChildren(cellElement, cellFormat, cellElements, cellCtx, false);
        FlushParagraph(cellElements, cellCtx);
        parentCtx.ImageIndex = cellCtx.ImageIndex;
        parentCtx.NextNumId = cellCtx.NextNumId;
        parentCtx.BulletAbstractNumId = cellCtx.BulletAbstractNumId;

        if (cellElements.Count == 0)
        {
            tc.Append(new Paragraph());
        }
        else
        {
            foreach (var el in cellElements)
            {
                tc.Append(el);
            }

            // OOXML requires every cell to end with a paragraph
            if (cellElements[^1] is not Paragraph)
            {
                tc.Append(new Paragraph());
            }
        }

        return tc;
    }

    static TableCellProperties ApplyCellStyles(Dictionary<string, string> declarations, TableCellProperties? tcPr)
    {
        tcPr ??= new();

        if (declarations.TryGetValue("width", out var cellWidth))
        {
            var twips = StyleParser.ParseLengthToTwips(cellWidth);
            if (twips != null)
            {
                tcPr.Append(new TableCellWidth
                {
                    Width = twips.Value.ToString(),
                    Type = TableWidthUnitValues.Dxa
                });
            }
        }

        if (declarations.TryGetValue("background-color", out var bgColor) ||
            declarations.TryGetValue("background", out bgColor))
        {
            var parsed = ColorParser.Parse(bgColor);
            if (parsed != null)
            {
                tcPr.Append(new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Fill = parsed
                });
            }
        }

        if (declarations.TryGetValue("vertical-align", out var vAlign))
        {
            var va = vAlign.AsSpan().Trim();
            var val = va.Equals("top", StringComparison.OrdinalIgnoreCase) ? TableVerticalAlignmentValues.Top
                : va.Equals("middle", StringComparison.OrdinalIgnoreCase) ? TableVerticalAlignmentValues.Center
                : va.Equals("bottom", StringComparison.OrdinalIgnoreCase) ? TableVerticalAlignmentValues.Bottom
                : (TableVerticalAlignmentValues?) null;
            if (val != null)
            {
                tcPr.Append(new TableCellVerticalAlignment
                {
                    Val = val.Value
                });
            }
        }

        if (declarations.TryGetValue("writing-mode", out var cellWritingMode))
        {
            var cwm = cellWritingMode.AsSpan().Trim();
            var cellTextDir = cwm.Equals("vertical-rl", StringComparison.OrdinalIgnoreCase) || cwm.Equals("tb-rl", StringComparison.OrdinalIgnoreCase)
                ? TextDirectionValues.TopToBottomRightToLeft
                : cwm.Equals("vertical-lr", StringComparison.OrdinalIgnoreCase) || cwm.Equals("tb-lr", StringComparison.OrdinalIgnoreCase)
                    ? TextDirectionValues.BottomToTopLeftToRight
                    : (TextDirectionValues?) null;
            if (cellTextDir != null)
            {
                tcPr.Append(new TextDirection
                {
                    Val = cellTextDir.Value
                });
            }
        }

        if (PaddingHelper.TryParse(declarations) is { } cellPad)
        {
            tcPr.Append(PaddingHelper.BuildMargin<TableCellMargin>(cellPad.Top, cellPad.Right, cellPad.Bottom, cellPad.Left));
        }

        BorderInfo? cellBorderAll = null;
        if (declarations.TryGetValue("border", out var cellBorderVal))
        {
            cellBorderAll = StyleParser.ParseBorder(cellBorderVal);
        }

        var cbt = declarations.TryGetValue("border-top", out var cbtVal) ? StyleParser.ParseBorder(cbtVal) : cellBorderAll;
        var cbr = declarations.TryGetValue("border-right", out var cbrVal) ? StyleParser.ParseBorder(cbrVal) : cellBorderAll;
        var cbb = declarations.TryGetValue("border-bottom", out var cbbVal) ? StyleParser.ParseBorder(cbbVal) : cellBorderAll;
        var cbl = declarations.TryGetValue("border-left", out var cblVal) ? StyleParser.ParseBorder(cblVal) : cellBorderAll;

        if (cbt != null || cbr != null || cbb != null || cbl != null)
        {
            var cellBorders = new TableCellBorders();
            BorderEmitter.AppendSides(cellBorders, cbt, cbl, cbb, cbr, 0);
            tcPr.Append(cellBorders);
        }

        return tcPr;
    }

    static void ApplyTableCellPadding(Dictionary<string, string> declarations, TableProperties tblPr)
    {
        if (PaddingHelper.TryParse(declarations) is { } pad)
        {
            tblPr.Append(PaddingHelper.BuildMargin<TableCellMarginDefault>(pad.Top, pad.Right, pad.Bottom, pad.Left));
        }
    }

    static List<int?> GetColumnWidths(IElement tableElement, int columnCount)
    {
        var widths = new List<int?>(columnCount);

        foreach (var child in tableElement.Children)
        {
            switch (child.LocalName)
            {
                case "colgroup":
                {
                    var hasColChild = false;
                    foreach (var gc in child.Children)
                    {
                        if (gc.LocalName == "col")
                        {
                            hasColChild = true;
                            AddColWidth(gc, widths);
                        }
                    }

                    if (!hasColChild)
                    {
                        AddColWidth(child, widths);
                    }

                    break;
                }
                case "col":
                    AddColWidth(child, widths);
                    break;
            }
        }

        while (widths.Count < columnCount)
        {
            widths.Add(null);
        }

        return widths;
    }

    static void AddColWidth(IElement col, List<int?> widths)
    {
        var span = ParseIntAttribute(col, "span", 1);
        var width = ParseColWidth(col);
        for (var i = 0; i < span; i++)
        {
            widths.Add(width);
        }
    }

    static int? ParseColWidth(IElement col)
    {
        var style = col.GetAttribute("style");
        if (style != null)
        {
            var declarations = StyleParser.Parse(style);
            if (declarations.TryGetValue("width", out var cssWidth))
            {
                var twips = StyleParser.ParseLengthToTwips(cssWidth);
                if (twips != null)
                {
                    return twips;
                }
            }
        }

        var widthAttr = col.GetAttribute("width");
        if (widthAttr != null && !widthAttr.EndsWith('%'))
        {
            return StyleParser.ParseLengthToTwips(widthAttr);
        }

        return null;
    }

    static List<IElement> GetTableRows(IElement tableElement)
    {
        var rows = new List<IElement>();
        foreach (var child in tableElement.Children)
        {
            switch (child.LocalName)
            {
                case "tr":
                    rows.Add(child);
                    break;
                case "thead" or "tbody" or "tfoot":
                {
                    foreach (var grandchild in child.Children)
                    {
                        if (grandchild.LocalName == "tr")
                        {
                            rows.Add(grandchild);
                        }
                    }

                    break;
                }
            }
        }

        return rows;
    }

    static List<IElement> GetCells(IElement row)
    {
        var cells = new List<IElement>();
        foreach (var child in row.Children)
        {
            if (child.LocalName is "td" or "th")
            {
                cells.Add(child);
            }
        }

        return cells;
    }

    static int GetColumnCount(List<IElement> rows)
    {
        var maxCols = 0;
        foreach (var row in rows)
        {
            var cols = 0;
            foreach (var cell in row.Children)
            {
                if (cell.LocalName is "td" or "th")
                {
                    cols += ParseIntAttribute(cell, "colspan", 1);
                }
            }

            maxCols = Math.Max(maxCols, cols);
        }

        return Math.Max(1, maxCols);
    }

    static int ParseIntAttribute(IElement element, string attribute, int defaultValue)
    {
        var value = element.GetAttribute(attribute);
        if (value != null && int.TryParse(value, out var result) && result > 1)
        {
            return result;
        }

        return defaultValue;
    }
}
