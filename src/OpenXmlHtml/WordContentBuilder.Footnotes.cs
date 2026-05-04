static partial class WordContentBuilder
{
    static Run BuildFootnoteRun(WordBuildContext context, string footnoteText)
    {
        context.FootnoteIndex++;
        var footnoteId = context.FootnoteIndex;

        var footnotesPart = context.MainPart!.FootnotesPart;
        if (footnotesPart == null)
        {
            footnotesPart = context.MainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new(
                new Footnote(
                    new Paragraph(
                        new Run(
                            new SeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.Separator,
                    Id = -1
                },
                new Footnote(
                    new Paragraph(
                        new Run(
                            new ContinuationSeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.ContinuationSeparator,
                    Id = 0
                });
        }

        footnotesPart.Footnotes!.Append(
            new Footnote(
                new Paragraph(
                    new Run(
                        new RunProperties(
                            new VerticalTextAlignment
                            {
                                Val = VerticalPositionValues.Superscript
                            }),
                        new FootnoteReferenceMark()),
                    new Run(
                        new Text(XmlCharFilter.StripInvalidXmlChars(" " + footnoteText))
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        })))
            {
                Id = footnoteId
            });

        return new(
            new RunProperties(
                new VerticalTextAlignment
                {
                    Val = VerticalPositionValues.Superscript
                }),
            new FootnoteReference
            {
                Id = footnoteId
            });
    }
}
