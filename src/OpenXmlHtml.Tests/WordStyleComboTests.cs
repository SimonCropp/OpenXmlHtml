[TestFixture]
public class WordStyleComboTests
{
    [Test]
    public Task FullFeatureDocx()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();

        // Add styles for CSS class mapping
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new Style(
                new StyleName { Val = "Subtle Reference" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(
                    new SpacingBetweenLines { After = "100" }))
            {
                StyleId = "SubtleReference",
                Type = StyleValues.Paragraph
            },
            new Style(
                new StyleName { Val = "Intense Emphasis" },
                new StyleRunProperties(
                    new WBold(),
                    new WItalic(),
                    new WColor { Val = "4472C4" }))
            {
                StyleId = "IntenseEmphasis",
                Type = StyleValues.Character
            });

        var body = new Body();
        mainPart.Document = new Document(body);

        WordHtmlConverter.AppendHtml(body,
            """
            <h1 style="text-align: center; margin-bottom: 24pt">Annual Report 2024</h1>
            <h2>Executive Summary</h2>
            <p style="text-indent: 36pt; line-height: 1.5; margin-bottom: 12pt">
              This report covers the <b>fiscal year 2024</b> performance across all divisions.
              Revenue grew by <span style="color: green; font-weight: bold">15%</span> year-over-year,
              driven primarily by our <i>cloud services</i> division.
            </p>

            <p class="SubtleReference">Source: Q4 2024 Financial Statements</p>

            <h2>Key Metrics</h2>
            <table>
              <caption>Quarterly Revenue (in millions)</caption>
              <thead>
                <tr><th>Quarter</th><th>Revenue</th><th>Growth</th></tr>
              </thead>
              <tbody>
                <tr>
                  <td>Q1</td>
                  <td style="text-align: right">$12.3M</td>
                  <td><span style="color: green">+8%</span></td>
                </tr>
                <tr>
                  <td>Q2</td>
                  <td style="text-align: right">$14.1M</td>
                  <td><span style="color: green">+12%</span></td>
                </tr>
                <tr>
                  <td>Q3</td>
                  <td style="text-align: right">$13.8M</td>
                  <td><span style="color: red">-2%</span></td>
                </tr>
                <tr>
                  <td>Q4</td>
                  <td style="text-align: right">$16.2M</td>
                  <td><span style="color: green">+17%</span></td>
                </tr>
              </tbody>
            </table>

            <h2>Strategic Priorities</h2>
            <ol>
              <li><b>Cloud expansion</b> — migrate remaining on-prem clients
                <ul>
                  <li>Target: <span style="color: #4472C4">95% cloud adoption</span> by Q2 2025</li>
                  <li>Investment: <code>$2.5M</code> in infrastructure</li>
                </ul>
              </li>
              <li><b>AI integration</b> — embed <abbr title="Machine Learning">ML</abbr> capabilities
                <ol>
                  <li>Natural language processing</li>
                  <li>Predictive analytics</li>
                </ol>
              </li>
              <li><del>Legacy system retirement</del> <ins>Completed in Q3</ins></li>
            </ol>

            <h3 style="margin-top: 24pt">Risk Factors</h3>
            <blockquote cite="https://example.com/risk-report">
              Market volatility and regulatory changes remain the <span class="IntenseEmphasis">primary concerns</span>
              for the upcoming fiscal year.
            </blockquote>

            <div style="margin: 18pt 36pt; text-align: justify; line-height: 150%">
              <p>This section uses combined CSS: justified text, custom margins on all sides,
              and 150% line spacing. It demonstrates that <u>multiple paragraph-level CSS properties</u>
              can be applied simultaneously alongside <b>inline formatting</b>.</p>
            </div>

            <h3>Team Contacts</h3>
            <dl>
              <dt>Engineering</dt>
              <dd>Contact <a href="mailto:eng@example.com">eng@example.com</a></dd>
              <dt>Sales</dt>
              <dd>Contact <a href="mailto:sales@example.com">sales@example.com</a></dd>
            </dl>

            <hr>

            <p style="text-align: center">
              <small>Confidential — <span style="font-family: Courier New">Internal Use Only</span></small>
            </p>

            <pre>Appendix A: Raw Data
            Q1: 12300000
            Q2: 14100000
            Q3: 13800000
            Q4: 16200000</pre>

            <p style="page-break-before: always; text-align: center; margin-top: 72pt">
              <sup>1</sup> All figures are in USD.
              See <a href="#appendix">appendix</a> for methodology.
            </p>

            <div id="appendix">
              <h2>Appendix: Methodology</h2>
              <p>Revenue figures are calculated using <q>accrual basis accounting</q>
              as defined by <abbr title="Generally Accepted Accounting Principles">GAAP</abbr>.</p>
            </div>
            """,
            mainPart);

        document.Dispose();
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
