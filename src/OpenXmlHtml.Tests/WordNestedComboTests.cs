[TestFixture]
public class WordNestedComboTests
{
    [Test]
    public Task NestedCombinations()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();
        mainPart.Document = new Document(body);

        WordHtmlConverter.AppendHtml(body,
            """
            <!-- Table inside table -->
            <h2>Nested Tables</h2>
            <table>
              <tr>
                <td>Outer cell 1</td>
                <td>
                  <table>
                    <tr><td>Inner A</td><td>Inner B</td></tr>
                    <tr><td>Inner C</td><td>Inner D</td></tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td>Outer cell 2</td>
                <td>Plain text</td>
              </tr>
            </table>

            <!-- List inside table -->
            <h2>List Inside Table</h2>
            <table>
              <tr>
                <th>Feature</th>
                <th>Details</th>
              </tr>
              <tr>
                <td>Languages</td>
                <td>
                  <ul>
                    <li>C#</li>
                    <li>F#</li>
                    <li>VB.NET</li>
                  </ul>
                </td>
              </tr>
              <tr>
                <td>Steps</td>
                <td>
                  <ol>
                    <li>Install SDK</li>
                    <li>Create project</li>
                    <li>Build and run</li>
                  </ol>
                </td>
              </tr>
            </table>

            <!-- Table inside list -->
            <h2>Table Inside List</h2>
            <ul>
              <li>Item with embedded table:
                <table>
                  <tr><th>Name</th><th>Value</th></tr>
                  <tr><td>Alpha</td><td>100</td></tr>
                </table>
              </li>
              <li>Normal item after table</li>
            </ul>

            <!-- Link inside bold/italic -->
            <h2>Links Inside Formatting</h2>
            <p><b><a href="https://example.com">Bold link</a></b></p>
            <p><i><a href="https://example.com/italic">Italic link</a></i></p>
            <p><b><i><a href="https://example.com/both">Bold italic link</a></i></b></p>
            <p><a href="https://example.com/mixed"><b>Bold</b> and <i>italic</i> inside link</a></p>

            <!-- Bold inside italic inside strikethrough -->
            <h2>Deeply Nested Formatting</h2>
            <p><s><i><b>Bold inside italic inside strikethrough</b></i></s></p>
            <p><u><sup>Superscript underlined</sup></u></p>
            <p><b><span style="color: red; font-size: 16pt">Bold red 16pt span</span></b></p>
            <p><mark><b><i>Bold italic highlighted</i></b></mark></p>
            <p><code><b>Bold code</b></code></p>

            <!-- Formatted content inside table cells -->
            <h2>Formatting Inside Table Cells</h2>
            <table>
              <tr>
                <td><b>Bold cell</b></td>
                <td><i>Italic cell</i></td>
                <td><a href="https://example.com">Link cell</a></td>
              </tr>
              <tr>
                <td><span style="color: green; font-weight: bold">Styled span</span></td>
                <td><code>Code cell</code></td>
                <td><mark>Highlighted</mark></td>
              </tr>
              <tr>
                <td><b><i><u>Bold italic underline</u></i></b></td>
                <td><sup>Super</sup>/<sub>sub</sub></td>
                <td><del>Deleted</del> <ins>Inserted</ins></td>
              </tr>
            </table>

            <!-- Nested lists (3 levels) -->
            <h2>Deeply Nested Lists</h2>
            <ul>
              <li>Level 1 bullet
                <ul>
                  <li>Level 2 bullet
                    <ul>
                      <li>Level 3 bullet</li>
                      <li><b>Bold</b> level 3</li>
                    </ul>
                  </li>
                </ul>
              </li>
            </ul>
            <ol>
              <li>Level 1 ordered
                <ol>
                  <li>Level 2 ordered
                    <ol>
                      <li>Level 3 ordered</li>
                    </ol>
                  </li>
                </ol>
              </li>
            </ol>

            <!-- Mixed list nesting -->
            <h2>Mixed List Types</h2>
            <ol>
              <li>Ordered item
                <ul>
                  <li>Unordered child
                    <ol>
                      <li>Back to ordered</li>
                    </ol>
                  </li>
                </ul>
              </li>
            </ol>

            <!-- Formatted content inside lists -->
            <h2>Formatting Inside Lists</h2>
            <ul>
              <li><b>Bold item</b> with <i>italic</i> and <code>code</code></li>
              <li><a href="https://example.com">Link item</a></li>
              <li><span style="color: blue; font-size: 14pt">Styled span item</span></li>
              <li>Item with <abbr title="abbreviation">abbr</abbr></li>
              <li><mark>Highlighted</mark> and <small>small</small></li>
            </ul>

            <!-- Blockquote with rich content -->
            <h2>Rich Blockquote</h2>
            <blockquote>
              <p><b>Important:</b> This blockquote contains <i>formatting</i>,
              <a href="https://example.com">links</a>, and <code>code</code>.</p>
              <ul>
                <li>Bullet inside blockquote</li>
                <li>Another bullet</li>
              </ul>
            </blockquote>

            <!-- Pre with inline formatting -->
            <h2>Preformatted With Formatting</h2>
            <pre>Line 1: plain
            Line 2: <b>bold in pre</b>
            Line 3: <span style="color: red">colored in pre</span></pre>

            <!-- Heading with inline formatting -->
            <h2><b>Bold</b> and <i>italic</i> in heading</h2>
            <h3><a href="https://example.com">Link inside heading</a></h3>
            <h3><code>Code</code> in heading</h3>

            <!-- Definition list with rich content -->
            <h2>Rich Definition List</h2>
            <dl>
              <dt><b>HTML</b></dt>
              <dd>A <i>markup</i> language. See <a href="https://example.com">docs</a>.</dd>
              <dt><code>CSS</code></dt>
              <dd>Style sheets for <span style="color: purple">visual presentation</span>.</dd>
            </dl>

            <!-- Paragraph-level CSS combos -->
            <h2>Paragraph Styling Combos</h2>
            <div style="margin: 12pt 24pt; background-color: #f0f0f0; border: 1px solid #ccc">
              <p style="text-indent: 36pt; text-align: justify; line-height: 1.8">
                Indented, justified, 1.8 line-height with <b>bold</b>, <i>italic</i>,
                and <span style="color: navy; font-family: Georgia">navy Georgia span</span>
                inside a bordered background div.
              </p>
            </div>

            <!-- Table with styling combos -->
            <h2>Styled Table</h2>
            <table style="border: 2px solid black; width: 100%">
              <tr>
                <td style="background-color: #eee; padding: 5pt; vertical-align: top; border: 1px solid gray">
                  <p style="text-align: center"><b>Centered bold</b></p>
                  <p><small>With small text below</small></p>
                </td>
                <td style="background-color: lightyellow; padding: 5pt; width: 60%">
                  <ol>
                    <li>List <b>inside</b> styled cell</li>
                    <li><a href="https://example.com">Link</a> in list in cell</li>
                  </ol>
                </td>
              </tr>
            </table>

            <!-- Nested table with formatting -->
            <h2>Nested Table With Formatting</h2>
            <table>
              <tr>
                <td>
                  <table style="border: 1px solid #999">
                    <tr>
                      <td><b><i>Bold italic</i></b></td>
                      <td><a href="https://example.com"><span style="color: red">Red link</span></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>

            <!-- Link wrapping formatted content -->
            <h2>Complex Link Content</h2>
            <p><a href="https://example.com"><b>Bold</b>, <i>italic</i>, and <code>code</code> in one link</a></p>
            <p><a href="https://example.com"><span style="font-size: 18pt; color: green">Large green link text</span></a></p>

            <!-- Inline quotation with formatting -->
            <h2>Quotation Combos</h2>
            <p>She said <q><b>hello</b> <i>world</i></q> emphatically.</p>
            <p><b>Bold with <q>quoted</q> text</b></p>

            <!-- Figure with formatted caption -->
            <h2>Figure With Rich Caption</h2>
            <figure>
              <figcaption><b>Figure 1:</b> A chart showing <i>growth</i> over <span style="color: blue">time</span></figcaption>
            </figure>

            <!-- HR between formatted sections -->
            <p><b>Section A</b></p>
            <hr>
            <p><i>Section B</i></p>

            <!-- Address with formatting -->
            <address>
              Contact <b>John Doe</b> at <a href="mailto:john@example.com">john@example.com</a>
            </address>

            <!-- Details/summary with formatting -->
            <details>
              <summary><b>Click to expand</b></summary>
              <p>Hidden content with <i>italic</i> and <a href="https://example.com">link</a>.</p>
            </details>
            """,
            mainPart);

        document.Dispose();
        stream.Position = 0;
        return Verify(stream, "docx");
    }
}
