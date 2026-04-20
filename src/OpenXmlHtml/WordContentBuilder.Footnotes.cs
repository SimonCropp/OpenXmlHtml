static partial class WordContentBuilder
{
    static Run BuildFootnoteRun(WordBuildContext ctx, string footnoteText)
    {
        ctx.FootnoteIndex++;
        var footnoteId = ctx.FootnoteIndex;

        var footnotesPart = ctx.MainPart!.FootnotesPart;
        if (footnotesPart == null)
        {
            footnotesPart = ctx.MainPart.AddNewPart<FootnotesPart>();
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
                        new Text(" " + footnoteText)
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
