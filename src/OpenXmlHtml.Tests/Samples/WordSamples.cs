[TestFixture]
public class WordSamples
{
    [Test]
    public Task ToParagraphs()
    {
        #region ToParagraphs

        var paragraphs = WordHtmlConverter.ToParagraphs(
            """
            <h1>Report Title</h1>
            <p>This is a <b>bold</b> statement with <i>emphasis</i>.</p>
            """);

        #endregion

        return Verify(paragraphs);
    }

    [Test]
    public Task AppendHtml()
    {
        #region AppendHtml

        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new(new Body());

        WordHtmlConverter.AppendHtml(
            main.Document.Body!,
            """
            <h1>Meeting Notes</h1>
            <p><i>Date: January 15, 2024</i></p>
            <ol>
              <li>Review <code>PR #123</code></li>
              <li>Update <u>documentation</u></li>
            </ol>
            """);

        #endregion

        return Verify(main.Document.Body!);
    }

    [Test]
    public Task RichDocument()
    {
        #region WordRichDocument

        var paragraphs = WordHtmlConverter.ToParagraphs(
            """
            <h2>Status Report</h2>
            <p>All systems <span style="color: green"><b>operational</b></span>.</p>
            <ul>
              <li>Server: <span style="color: green">OK</span></li>
              <li>Cache: <span style="color: red">Down</span></li>
            </ul>
            <p>Contact <a href="mailto:ops@example.com">ops team</a> for details.</p>
            """);

        #endregion

        return Verify(paragraphs);
    }

    [Test]
    public Task ConvertToDocx()
    {
        #region ConvertToDocx

        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <h1>Report</h1>
            <p>This is a <b>bold</b> statement.</p>
            <ul>
              <li>Item one</li>
              <li>Item two</li>
            </ul>
            """,
            stream);

        #endregion

        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task ConvertStreamToDocx()
    {
        #region ConvertStreamToDocx

        using var htmlStream = new MemoryStream(
            "<h1>Report</h1><p>Content</p>"u8.ToArray());
        using var docxStream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(htmlStream, docxStream);

        #endregion

        docxStream.Position = 0;
        return Verify(docxStream, "docx");
    }

    [Test]
    public Task EmployeeOnboardingGuide()
    {
        var logo = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";
        using var stream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            $"""
            <p><img src="data:image/png;base64,{logo}" width="64" height="64"></p>
            <h1>Employee Onboarding Guide</h1>
            <p><i>Human Resources Department</i></p>
            <hr>
            <h2>Welcome to <span style="color: #0563C1">Contoso Ltd</span></h2>
            <p>We're excited to have you join the team! Below you'll find everything
            you need to get started on your <b>first day</b>.</p>

            <h3>Checklist</h3>
            <ol>
              <li>Sign employment agreement</li>
              <li>Complete <a href="https://hr.contoso.com/tax">tax forms</a></li>
              <li>Collect your <mark>badge and laptop</mark> from IT</li>
              <li>Read the <u>code of conduct</u></li>
            </ol>

            <h3>Key Contacts</h3>
            <table>
              <caption>Department Contacts</caption>
              <thead>
                <tr><th>Department</th><th>Contact</th><th>Extension</th></tr>
              </thead>
              <tbody>
                <tr><td>IT Support</td><td>helpdesk@contoso.com</td><td><code>x4100</code></td></tr>
                <tr><td>Facilities</td><td>facilities@contoso.com</td><td><code>x4200</code></td></tr>
                <tr><td>HR</td><td>hr@contoso.com</td><td><code>x4300</code></td></tr>
              </tbody>
            </table>

            <h3>Important Policies</h3>
            <dl>
              <dt>Remote Work</dt>
              <dd>Up to <b>3 days per week</b> after probation period.</dd>
              <dt>Time Off</dt>
              <dd>20 days PTO plus <span style="color: green">10 public holidays</span>.</dd>
            </dl>

            <blockquote>
              <q>The strength of the team is each individual member.</q>
              — Phil Jackson
            </blockquote>

            <p><small>Last updated: January 2024. For questions contact
            <a href="mailto:onboarding@contoso.com">onboarding@contoso.com</a>.</small></p>
            """,
            stream);
        stream.Position = 0;
        return Verify(stream, "docx");
    }

    [Test]
    public Task RemoteImageSettings()
    {
        #region RemoteImageSettings

        // Configure which image sources are allowed
        var settings = new HtmlConvertSettings
        {
            WebImages = ImagePolicy.SafeDomains("cdn.example.com"),
            LocalImages = ImagePolicy.SafeDirectories(@"C:\Reports\Images")
        };

        // Pass settings to any conversion method
        using var settingsStream = new MemoryStream();
        WordHtmlConverter.ConvertToDocx(
            """
            <h1>Report</h1>
            <p><img src="https://cdn.example.com/logo.png" alt="Logo"></p>
            """,
            settingsStream,
            settings);

        #endregion

        settingsStream.Position = 0;
        return Verify(settingsStream, "docx");
    }

    [Test]
    public void SharedNumberingSession()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new(new Body());

        #region SharedNumberingSession

        var session = new HtmlNumberingSession();
        var settings = new HtmlConvertSettings
        {
            NumberingSession = session
        };

        // Both fragments reuse one bullet abstract — one definition, two numbering instances.
        var first = WordHtmlConverter.ToElements("<ul><li>a</li><li>b</li></ul>", mainPart, settings);
        var second = WordHtmlConverter.ToElements("<ul><li>c</li><li>d</li></ul>", mainPart, settings);

        #endregion

        _ = first;
        _ = second;
    }

    [Test]
    public Task HeadersAndFooters()
    {
        #region HeadersAndFooters

        using var headerStream = new MemoryStream();
        using var headerDoc = WordprocessingDocument.Create(
            headerStream, WordprocessingDocumentType.Document);
        var headerMainPart = headerDoc.AddMainDocumentPart();
        headerMainPart.Document = new(new Body());

        WordHtmlConverter.AppendHtml(
            headerMainPart.Document.Body!,
            "<p>Document content</p>",
            headerMainPart);

        WordHtmlConverter.SetHeader(headerMainPart,
            """<p style="text-align: center"><b>Company Name</b></p>""");

        WordHtmlConverter.SetFooter(headerMainPart,
            """<p style="text-align: center; font-size: 9pt; color: gray">Confidential</p>""");

        #endregion

        headerDoc.Dispose();
        headerStream.Position = 0;
        return Verify(headerStream, "docx");
    }

    [Test]
    public Task StyleMapping()
    {
        using var styleStream = new MemoryStream();
        using var styleDoc = WordprocessingDocument.Create(
            styleStream, WordprocessingDocumentType.Document);
        var styleMainPart = styleDoc.AddMainDocumentPart();

        var stylesPart = styleMainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new(
            new Style
            {
                StyleId = "Quote",
                Type = StyleValues.Paragraph
            },
            new Style
            {
                StyleId = "Emphasis",
                Type = StyleValues.Character
            });

        var styleBody = new Body();
        styleMainPart.Document = new(styleBody);

        #region StyleMapping

        // CSS class names are matched against Word styles in the document.
        // Paragraph styles apply via ParagraphStyleId,
        // character styles apply via RunStyle.
        WordHtmlConverter.AppendHtml(
            styleBody,
            """
            <p class="Quote">This uses the Quote paragraph style</p>
            <p>Normal text with <span class="Emphasis">emphasized</span> word</p>
            """,
            styleMainPart);

        #endregion

        styleDoc.Dispose();
        styleStream.Position = 0;
        return Verify(styleStream, "docx");
    }

    [Test]
    public async Task ConvertFileToDocx()
    {
        var htmlPath = await TempFile.CreateText("<h1>Hello</h1><p>World</p>");
        var docxPath = new TempFile("docx");

        #region ConvertFileToDocx

        WordHtmlConverter.ConvertFileToDocx(htmlPath, docxPath);

        #endregion

        await VerifyFile(docxPath);
    }
}
