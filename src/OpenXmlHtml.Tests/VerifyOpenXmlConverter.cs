static class VerifyOpenXmlConverter
{
    internal static void Initialize()
    {
        VerifierSettings.RegisterFileConverter<SpreadsheetInlineString>(ConvertInlineString);
        VerifierSettings.RegisterFileConverter<SpreadsheetCell>(ConvertCell);
        VerifierSettings.RegisterFileConverter<List<Paragraph>>(ConvertParagraphs);
        VerifierSettings.RegisterFileConverter<Body>(ConvertBody);
    }

    static ConversionResult ConvertInlineString(SpreadsheetInlineString value, IReadOnlyDictionary<string, object> context) =>
        new(null, "xml", value.OuterXml);

    static ConversionResult ConvertCell(SpreadsheetCell value, IReadOnlyDictionary<string, object> context) =>
        new(null, "xml", value.OuterXml);

    static ConversionResult ConvertParagraphs(List<Paragraph> value, IReadOnlyDictionary<string, object> context) =>
        new(null, "xml", string.Join('\n', value.Select(_ => _.OuterXml)));

    static ConversionResult ConvertBody(Body value, IReadOnlyDictionary<string, object> context) =>
        new(null, "xml", value.OuterXml);
}
