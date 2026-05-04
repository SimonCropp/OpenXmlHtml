static class BorderEmitter
{
    static T Side<T>(BorderInfo border, byte space)
        where T : BorderType, new() =>
        new()
        {
            Val = border.Style,
            Size = (uint) border.SizeEighths,
            Space = space,
            Color = border.Color ?? "auto"
        };

    static T NoneSide<T>()
        where T : BorderType, new() =>
        new()
        {
            Val = BorderValues.None,
            Size = 0,
            Space = 0
        };

    internal static TableBorders BuildTableBorders(BorderInfo border)
    {
        if (border.Style == BorderValues.None)
        {
            return new(
                NoneSide<TopBorder>(),
                NoneSide<LeftBorder>(),
                NoneSide<BottomBorder>(),
                NoneSide<RightBorder>(),
                NoneSide<InsideHorizontalBorder>(),
                NoneSide<InsideVerticalBorder>());
        }

        return new(
            Side<TopBorder>(border, 0),
            Side<LeftBorder>(border, 0),
            Side<BottomBorder>(border, 0),
            Side<RightBorder>(border, 0),
            Side<InsideHorizontalBorder>(border, 0),
            Side<InsideVerticalBorder>(border, 0));
    }

    internal static void AppendSides(
        OpenXmlCompositeElement container,
        BorderInfo? top,
        BorderInfo? left,
        BorderInfo? bottom,
        BorderInfo? right,
        byte space)
    {
        if (top != null && top.Style != BorderValues.None)
        {
            container.Append(Side<TopBorder>(top, space));
        }

        if (left != null && left.Style != BorderValues.None)
        {
            container.Append(Side<LeftBorder>(left, space));
        }

        if (bottom != null && bottom.Style != BorderValues.None)
        {
            container.Append(Side<BottomBorder>(bottom, space));
        }

        if (right != null && right.Style != BorderValues.None)
        {
            container.Append(Side<RightBorder>(right, space));
        }
    }
}
