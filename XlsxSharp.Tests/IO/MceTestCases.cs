#nullable enable
using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

internal record MceTestCase(
    string Description,
    string Xml,
    HashSet<string> AppConfig,
    List<IExpectedXmlNode> ExpectedNodes,
    (string LocalName, string NamespaceUri)? Adee = null
)
{
    private const string MceNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string FooNs = "http://www.example.com/foo";
    private const string BarNs = "http://www.example.com/bar";
    private const string ExampleNs = "http://www.example.com";

    public static IEnumerable<MceTestCase> ProcessingTestCases()
    {
        // XML from semantic example of processing in chapter 9.4
        const string semanticExample = $"""
            <example xmlns="{ExampleNs}"
                     xmlns:mce="{MceNs}"
                     xmlns:foo="{FooNs}"
                     xmlns:bar="{BarNs}"
                     mce:Ignorable="foo bar"
                     mce:ProcessContent="foo:unwrapped">
              <mce:AlternateContent>
                <mce:Choice Requires="foo"> <!-- Choice #1 -->
                  <foo:foo/> <!-- Foo #1 -->
                  <bar:bar> <!-- Bar #1 -->
                    <mce:AlternateContent>
                      <mce:Choice Requires="bar"> <!-- Choice #1-1 -->
                        <Choice1-1/>
                      </mce:Choice>
                      <mce:Fallback> <!-- Fallback #1-1 -->
                        <Fallback1-1/>
                      </mce:Fallback>
                    </mce:AlternateContent>
                  </bar:bar>
                </mce:Choice>
                <mce:Choice Requires="bar"> <!-- Choice #2 -->
                  <bar:bar/> <!-- Bar #2 -->
                  <foo:unwrapped> <!-- Foo #2 -->
                    <mce:AlternateContent>
                      <mce:Choice Requires="foo"> <!-- Choice #2-1 -->
                        <Choice2-1/>
                      </mce:Choice>
                      <mce:Fallback> <!-- Fallback #2-1 -->
                        <Fallback2-1/>
                      </mce:Fallback>
                    </mce:AlternateContent>
                  </foo:unwrapped>
                </mce:Choice>
              </mce:AlternateContent>
            </example>
            """;

        yield return new MceTestCase(
            "9.4: Processing model example - only foo NS",
            semanticExample,
            [FooNs],
            [
                new Open("example", ExampleNs),
                new Open("foo", FooNs),
                new Close("foo", FooNs),
                new Close("example", ExampleNs),
            ]
        );

        yield return new MceTestCase(
            "9.4: Processing model example - only bar NS",
            semanticExample,
            [BarNs],
            [
                new Open("example", ExampleNs),
                new Open("bar", BarNs),
                new Close("bar", BarNs),
                new Open("Fallback2-1", ExampleNs),
                new Close("Fallback2-1", ExampleNs),
                new Close("example", ExampleNs),
            ]
        );

        yield return new MceTestCase(
            "9.4: Processing model example - both foo and bar NS",
            semanticExample,
            [FooNs, BarNs],
            [
                new Open("example", ExampleNs),
                new Open("foo", FooNs),
                new Close("foo", FooNs),
                new Open("bar", BarNs),
                new Open("Choice1-1", ExampleNs),
                new Close("Choice1-1", ExampleNs),
                new Close("bar", BarNs),
                new Close("example", ExampleNs),
            ]
        );

        // Example from chapter A.2.3: Processable content
        const string circles1Ns = "http://www.example.com/Circles/v1";
        const string circles2Ns = "http://www.example.com/Circles/v2";
        const string circles3Ns = "http://www.example.com/Circles/v3";
        const string processableContentExample = $"""
            <Circles xmlns="{circles1Ns}"
                     xmlns:v2="{circles2Ns}"
                     xmlns:mc="{MceNs}"
                     mc:Ignorable="v2"
                     mc:ProcessContent="v2:Blink">
              <v2:Watermark Opacity="v0.1">
                <Circle Center="0,0" Radius="20" Color="Blue"/>
              </v2:Watermark>
              <v2:Blink>
                <Circle Center="13,0" Radius="20" Color="Yellow"/>
              </v2:Blink>
            </Circles>
            """;

        yield return new MceTestCase(
            "A2.3.: Processable content example - 1",
            processableContentExample,
            [circles1Ns, circles2Ns],
            [
                new Open("Circles", circles1Ns),
                new Open("Watermark", circles2Ns),
                new Open("Circle", circles1Ns).WithAttribute("Color", "Blue"),
                new Close("Circle", circles1Ns),
                new Close("Watermark", circles2Ns),
                new Open("Blink", circles2Ns),
                new Open("Circle", circles1Ns).WithAttribute("Color", "Yellow"),
                new Close("Circle", circles1Ns),
                new Close("Blink", circles2Ns),
                new Close("Circles", circles1Ns),
            ]
        );

        yield return new MceTestCase(
            "A2.3.: Processable content example - 2",
            processableContentExample,
            [circles1Ns],
            [
                new Open("Circles", circles1Ns),
                new Open("Circle", circles1Ns).WithAttribute("Color", "Yellow"),
                new Close("Circle", circles1Ns),
                new Close("Circles", circles1Ns),
            ]
        );

        // Example from chapter A.2.6: AlternateContent Element
        const string acExample = $"""
            <Circles xmlns="{circles1Ns}"
                     xmlns:v2="{circles2Ns}"
                     xmlns:v3="{circles3Ns}"
                     xmlns:mc="{MceNs}"
                     mc:Ignorable="v2 v3">
              <mc:AlternateContent>
                <mc:Choice Requires="v3">
                  <v3:Circle Center="0,0" Radius="20" Color="Blue" Opacity="0.5" Luminance="13"/>
                </mc:Choice>
                <mc:Fallback>
                  <LuminanceFilter Luminance="13">
                    <Circle Center="0,0" Radius="20" Color="Blue" v2:Opacity="0.5"/>
                  </LuminanceFilter>
                </mc:Fallback>
              </mc:AlternateContent>
            </Circles>
            """;

        yield return new MceTestCase(
            "A.2.3: AlternateContent example - 1",
            acExample,
            [circles1Ns, circles2Ns, circles3Ns],
            [
                new Open("Circles", circles1Ns),
                new Open("Circle", circles3Ns).WithAttribute("Luminance", "13"),
                new Close("Circle", circles3Ns),
                new Close("Circles", circles1Ns),
            ]
        );

        yield return new MceTestCase(
            "A.2.3: AlternateContent example - 2",
            acExample,
            [circles1Ns, circles2Ns],
            [
                new Open("Circles", circles1Ns),
                new Open("LuminanceFilter", circles1Ns),
                new Open("Circle", circles1Ns).WithAttribute("Color", "Blue"),
                new Close("Circle", circles1Ns),
                new Close("LuminanceFilter", circles1Ns),
                new Close("Circles", circles1Ns),
            ]
        );

        // Can emit even text in a mixed content and doesn't emit the text inside the AC, but outside of a choice
        yield return new MceTestCase(
            "AC: Omit text elements outside of selected choice",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              Prefix
              <mc:AlternateContent>
                Not Emitted 1
                <mc:Choice Requires="foo">
                  Infix
                </mc:Choice>
                Not Emitted 2
              </mc:AlternateContent>
              Postfix
            </root>
            """,
            [FooNs],
            [
                new Open("root"),
                new Text("Prefix"),
                new Text("Infix"),
                new Text("Postfix"),
                new Close("root"),
            ]
        );

        // AC can contain other elements from ignorable namespaces for future-proofing
        yield return new MceTestCase(
            "AC: Allow ignorable elements as children for future use",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" xmlns:bar="{BarNs}" mc:Ignorable="bar">
              <mc:AlternateContent>
                <bar:FutureChoice/>
                <mc:Choice Requires="foo">
                  Selected choice
                </mc:Choice>
                <mc:Fallback>
                  Fallback
                </mc:Fallback>
                <bar:FutureFallback>
                  Future fallback
                </bar:FutureFallback>
              </mc:AlternateContent>
            </root>
            """,
            [FooNs],
            [new Open("root"), new Text("Selected choice"), new Close("root")]
        );

        // Choice will take fallbackAC inside another AC
        yield return new MceTestCase(
            "AC: If no choice selected, use fallback",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:AlternateContent>
                <mc:Choice Requires="foo">
                  Choice
                </mc:Choice>
                <mc:Fallback>
                  Fallback
                </mc:Fallback>
              </mc:AlternateContent>
            </root>
            """,
            [],
            [new Open("root"), new Text("Fallback"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: AC without selected choice and no fallback is completely omitted",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:AlternateContent>
                <mc:Choice Requires="foo">Choice</mc:Choice>
              </mc:AlternateContent>
            </root>
            """,
            [],
            [new Open("root"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: First choice with its required namespaces in app config is selected",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" xmlns:bar="{BarNs}">
              <mc:AlternateContent>
                <mc:Choice Requires="foo">Choice foo</mc:Choice>
                <mc:Choice Requires="bar">Choice bar</mc:Choice>
                <mc:Choice Requires="foo bar">Choice foo bar</mc:Choice>
                <mc:Fallback>Fallback</mc:Fallback>
              </mc:AlternateContent>
            </root>
            """,
            [FooNs, BarNs],
            [new Open("root"), new Text("Choice foo"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: ACs can be nested",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" xmlns:bar="{BarNs}">
              <mc:AlternateContent>
                <mc:Choice Requires="foo">Choice 1-1</mc:Choice>
                <mc:Choice Requires="bar">
                  <mc:AlternateContent>
                    <mc:Choice Requires="foo">Choice 2-1</mc:Choice>
                    <mc:Fallback>Fallback 2-2</mc:Fallback>
                  </mc:AlternateContent>
                </mc:Choice>
              </mc:AlternateContent>
            </root>
            """,
            [BarNs],
            [new Open("root"), new Text("Fallback 2-2"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: ACs can be in a sequence and no selected",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:AlternateContent>
                <mc:Choice Requires="foo">Choice 1</mc:Choice>
              </mc:AlternateContent>
              <mc:AlternateContent>
                <mc:Choice Requires="foo">Choice 2</mc:Choice>
              </mc:AlternateContent>
            </root>
            """,
            [],
            [new Open("root"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: AC can in root",
            $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:Choice Requires="foo">
                <root/>
              </mc:Choice>
            </mc:AlternateContent>
            """,
            [FooNs],
            [new Open("root"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: AC in root can be empty",
            $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:Choice Requires="foo">
                <root/>
              </mc:Choice>
            </mc:AlternateContent>
            """,
            [],
            []
        );

        yield return new MceTestCase(
            "AC: AC self-closing choice is valid",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:AlternateContent >
                <mc:Choice Requires="foo"/>
              </mc:AlternateContent>
            </root>
            """,
            [FooNs],
            [new Open("root"), new Close("root")]
        );

        yield return new MceTestCase(
            "AC: AC choice with only insignificant whitespace inside is valid",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:AlternateContent >
                <mc:Choice Requires="foo">   </mc:Choice>
              </mc:AlternateContent>
            </root>
            """,
            [FooNs],
            [new Open("root"), new Close("root")]
        );

        yield return new MceTestCase(
            "PC: PC with star unwraps every element",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:ProcessContent="foo:*" mc:Ignorable="foo">
              <foo:unwrapFirst>
                <first/>
              </foo:unwrapFirst>
              <foo:unwrapSecond>
                <second/>
              </foo:unwrapSecond>
            </root>
            """,
            [],
            [
                new Open("root"),
                new Open("first"),
                new Close("first"),
                new Open("second"),
                new Close("second"),
                new Close("root"),
            ]
        );

        yield return new MceTestCase(
            "PC: PC with specific name unwraps only that element and ignores others",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:ProcessContent="foo:unwrapFirst" mc:Ignorable="foo">
              <foo:unwrapFirst>
                <first/>
              </foo:unwrapFirst>
              <!-- Any process content namespace must be declared as ignorable and if not part of of app config, it is ignored -->
              <foo:unwrapSecond>
                <second/>
              </foo:unwrapSecond>
            </root>
            """,
            [],
            [new Open("root"), new Open("first"), new Close("first"), new Close("root")]
        );

        yield return new MceTestCase(
            "PC: PC with with namespace that is part of app config is not unwrapped or ignored",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:ProcessContent="foo:*" mc:Ignorable="foo">
              <foo:processContent>
                <item/>
              </foo:processContent>
            </root>
            """,
            [FooNs],
            [
                new Open("root"),
                new Open("processContent", FooNs),
                new Open("item"),
                new Close("item"),
                new Close("processContent", FooNs),
                new Close("root"),
            ]
        );

        yield return new MceTestCase(
            "Ignore: Ignorable namespaces (foo) that are not part of appConfig are skipped",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:Ignorable="foo">
              <foo:element>
                <skippedItem/>
              </foo:element>
              <item/>
              <foo:selfClosing/>
            </root>
            """,
            [],
            [new Open("root"), new Open("item"), new Close("item"), new Close("root")]
        );

        yield return new MceTestCase(
            "Ignore: Ignorable namespaces (foo) that are part of appConfig are emitted",
            $"""
            <root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:Ignorable="foo">
              <foo:element>
                <item/>
              </foo:element>
              <item/>
              <foo:selfClosing/>
            </root>
            """,
            [FooNs],
            [
                new Open("root"),
                new Open("element", FooNs),
                new Open("item"),
                new Close("item"),
                new Close("element", FooNs),
                new Open("item"),
                new Close("item"),
                new Open("selfClosing", FooNs),
                new Close("selfClosing", FooNs),
                new Close("root"),
            ]
        );

        yield return new MceTestCase(
            "ADEE: Empty ADEE is emitted",
            $"""
            <root xmlns:mc="{MceNs}" xmlns="{FooNs}">
              <ext/>
            </root>
            """,
            [FooNs],
            [
                new Open("root", FooNs),
                new Open("ext", FooNs),
                new Close("ext", FooNs),
                new Close("root", FooNs),
            ],
            ("ext", FooNs)
        );

        yield return new MceTestCase(
            "ADEE: ADEE emits even MCE elements",
            $"""
            <root xmlns:mc="{MceNs}" xmlns="{FooNs}" xmlns:bar="{BarNs}">
              <ext>
                <mc:AlternateContent>
                  <mc:Choice Requires="foo">Choice1-1</mc:Choice>
                  <mc:Fallback>Fallback</mc:Fallback>
                </mc:AlternateContent>
                <insideAdee/>
              </ext>
              <mc:AlternateContent>
                <mc:Choice Requires="bar">
                  <bar:bar/>
                </mc:Choice>
                <mc:Fallback>Fallback</mc:Fallback>
              </mc:AlternateContent>
            </root>
            """,
            [FooNs, BarNs],
            [
                new Open("root", FooNs),
                new Open("ext", FooNs),
                new Open("AlternateContent", MceNs),
                new Open("Choice", MceNs),
                new Text("Choice1-1"),
                new Close("Choice", MceNs),
                new Open("Fallback", MceNs),
                new Text("Fallback"),
                new Close("Fallback", MceNs),
                new Close("AlternateContent", MceNs),
                new Open("insideAdee", FooNs),
                new Close("insideAdee", FooNs),
                new Close("ext", FooNs),
                new Open("bar", BarNs),
                new Close("bar", BarNs),
                new Close("root", FooNs),
            ],
            ("ext", FooNs)
        );
    }

    private record Open(
        string LocalName,
        string? NamespaceUri = null,
        Dictionary<string, string>? Attributes = null
    ) : IExpectedXmlNode
    {
        public Open WithAttribute(string name, string value) =>
            this with
            {
                Attributes = new Dictionary<string, string>(this.Attributes ?? [])
                {
                    { name, value },
                },
            };

        public void AssertMatches(IXmlReader reader)
        {
            ClassicAssert.AreEqual(XmlTreeNodeType.OpenElement, reader.NodeType);
            ClassicAssert.AreEqual(this.LocalName, reader.LocalName);
            ClassicAssert.AreEqual(this.NamespaceUri ?? string.Empty, reader.NamespaceUri);
            ClassicAssert.IsEmpty(reader.Value);

            if (this.Attributes is not null)
            {
                foreach ((string name, string expectedAttributeValue) in this.Attributes)
                {
                    string? attributeValue = reader.GetAttribute(name, null);
                    ClassicAssert.AreEqual(expectedAttributeValue, attributeValue);
                }
            }
        }
    }

    private record Close(string LocalName, string? NamespaceUri = null) : IExpectedXmlNode
    {
        public void AssertMatches(IXmlReader reader)
        {
            ClassicAssert.AreEqual(XmlTreeNodeType.CloseElement, reader.NodeType);
            ClassicAssert.AreEqual(this.LocalName, reader.LocalName);
            ClassicAssert.AreEqual(this.NamespaceUri ?? string.Empty, reader.NamespaceUri);
            ClassicAssert.IsEmpty(reader.Value);
        }
    }

    private record Text(string Value, bool Trim = true) : IExpectedXmlNode
    {
        public void AssertMatches(IXmlReader reader)
        {
            ClassicAssert.AreEqual(XmlTreeNodeType.Text, reader.NodeType);
            ClassicAssert.AreEqual(this.Value, this.Trim ? reader.Value.Trim() : reader.Value);
            ClassicAssert.IsEmpty(reader.LocalName);
            ClassicAssert.IsEmpty(reader.NamespaceUri);
        }
    }

    public override string ToString() => this.Description;
}
