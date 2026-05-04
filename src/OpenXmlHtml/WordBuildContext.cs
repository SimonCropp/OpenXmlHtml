class WordBuildContext
{
    internal List<OpenXmlElement> CurrentRuns = [];
    internal int ImageIndex;
    internal int ListDepth;
    internal int HeadingLevel;
    internal int FootnoteIndex;
    internal int BookmarkId;
    internal MainDocumentPart? MainPart;
    internal HtmlConvertSettings? Settings;
    internal Dictionary<string, StyleType>? StyleMap;
    internal string? ParagraphStyleId;
    internal ParagraphFormatState? ParagraphFormat;
    internal Stack<(int NumId, int Ilvl, bool IsOrdered)> ListStack = new();
    internal int? BulletAbstractNumId;
    internal int NextNumId;
    internal int? ListNumId;
    internal int? ListIlvl;
    internal int? ReversedStart;
}
