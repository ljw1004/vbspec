Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports FSharp.Markdown
Imports Microsoft.FSharp.Core


Module Module1

    Sub Main()
        ' Pick a output filename that's not already locked by Word...
        Dim fn = ""
        For ifn = 0 To 1000
            fn = $"vbmd{If(ifn = 0, "", CStr(ifn))}.docx"
            Try
                If File.Exists(fn) Then File.Delete(fn) Else Exit For
            Catch ex As Exception
            End Try
        Next

        Using templateDoc = WordprocessingDocument.Open("vb-template.docx", False),
                resultDoc = WordprocessingDocument.Create(fn, WordprocessingDocumentType.Document)
            For Each part In templateDoc.Parts
                resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId)
            Next

            Dim body = resultDoc.MainDocumentPart.Document.Body
            body.RemoveAllChildren()

            Dim src = "> __Note__
> This is about `code` spacing"
            src = ""

            For Each p In New MarkdownParser(src, resultDoc).Paragraphs
                body.AppendChild(p)
            Next

            If String.IsNullOrEmpty(src) Then
                For Each mdfn In {"introduction.md", "lexical-grammar.md", "preprocessing-directives.md",
                "general-concepts.md", "attributes.md", "source-files-and-namespaces.md", "types.md",
                "conversions.md"}
                    For Each p In New MarkdownParser(File.ReadAllText($"..\..\..\vb\{mdfn}"), resultDoc).Paragraphs
                        body.AppendChild(p)
                    Next
                Next
            End If

        End Using

        Process.Start(fn)
    End Sub

    <Extension>
    Public Function [Option](Of T As Class)(this As FSharpOption(Of T)) As T
        If FSharpOption(Of T).GetTag(this) = FSharpOption(Of T).Tags.None Then Return Nothing
        Return this.Value
    End Function

End Module

Class MarkdownParser
    Dim mddoc As MarkdownDocument
    Dim wdoc As WordprocessingDocument

    Sub New(s As String, wdoc As WordprocessingDocument)
        Me.mddoc = Markdown.Parse(s)
        Me.wdoc = wdoc
    End Sub

    Function Paragraphs() As IEnumerable(Of OpenXmlCompositeElement)
        Return Paragraphs2Paragraphs(mddoc.Paragraphs)
    End Function

    Iterator Function Paragraphs2Paragraphs(pars As IEnumerable(Of MarkdownParagraph)) As IEnumerable(Of OpenXmlCompositeElement)
        For Each md In pars
            For Each p In Paragraph2Paragraphs(md)
                Yield p
            Next
        Next
    End Function

    Dim re_grammar As New Regex("^\s*([a-zA-Z]+):\s")

    Iterator Function Paragraph2Paragraphs(md As MarkdownParagraph) As IEnumerable(Of OpenXmlCompositeElement)
        If md.IsHeading Then
            Dim mdh = CType(md, MarkdownParagraph.Heading), level = mdh.Item1, spans = mdh.Item2
            Dim props As New ParagraphProperties(New ParagraphStyleId() With {.Val = $"Heading{level}"})
            Yield New Paragraph(Spans2Elements(spans)) With {.ParagraphProperties = props}
            Return

        ElseIf md.IsParagraph Then
            Dim mdp = CType(md, MarkdownParagraph.Paragraph), spans = mdp.Item
            Yield New Paragraph(Spans2Elements(spans))
            Return

        ElseIf md.IsListBlock Then
            Dim mdl = CType(md, MarkdownParagraph.ListBlock)
            Dim flat = FlattenList(mdl)

            ' Let's figure out what kind of list it is - ordered or unordered? nested?
            Dim format0 = {"1", "1", "1"}
            For Each item In flat
                format0(item.Item1) = If(item.Item2, "1", "o")
            Next
            Dim format = String.Join("", format0)

            Dim numberingPart = If(wdoc.MainDocumentPart.NumberingDefinitionsPart, wdoc.MainDocumentPart.AddNewPart(Of NumberingDefinitionsPart)("NumberingDefinitionsPart001"))
            If numberingPart.Numbering Is Nothing Then numberingPart.Numbering = New Numbering()

            Dim createLevel = Function(level As Integer, isOrdered As Boolean) As Level
                                  Dim numFormat = NumberFormatValues.Bullet, levelText = "·"
                                  If isOrdered AndAlso level = 0 Then numFormat = NumberFormatValues.Decimal : levelText = "%1."
                                  If isOrdered AndAlso level = 1 Then numFormat = NumberFormatValues.LowerLetter : levelText = "%2."
                                  If isOrdered AndAlso level = 2 Then numFormat = NumberFormatValues.LowerRoman : levelText = "%3."
                                  Return New Level(New StartNumberingValue With {.Val = 1}, New NumberingFormat With {.Val = numFormat}, New LevelText With {.Val = levelText}, New ParagraphProperties(New Indentation With {.Left = CStr(540 + 360 * level), .Hanging = "360"})) With {.LevelIndex = level}
                              End Function
            Dim level0 = createLevel(0, format(0) = "1")
            Dim level1 = createLevel(1, format(1) = "1")
            Dim level2 = createLevel(2, format(2) = "1")

            Dim abstracts = numberingPart.Numbering.OfType(Of AbstractNum).Select(Function(an) an.AbstractNumberId.Value).ToList()
            Dim aid = If(abstracts.Count = 0, 1, abstracts.Max() + 1)
            Dim abstract As New AbstractNum(New MultiLevelType() With {.Val = MultiLevelValues.Multilevel}, level0, level1, level2) With {.AbstractNumberId = aid}
            numberingPart.Numbering.InsertAt(abstract, 0)

            Dim instances = numberingPart.Numbering.OfType(Of NumberingInstance).Select(Function(ni) ni.NumberID.Value)
            Dim nid = If(instances.Count = 0, 1, instances.Max() + 1)
            Dim numInstance As New NumberingInstance(New AbstractNumId With {.Val = aid}) With {.NumberID = nid}
            numberingPart.Numbering.AppendChild(numInstance)

            For Each item In flat
                Dim spans = item.Item3.Item
                Yield New Paragraph(Spans2Elements(spans)) With {.ParagraphProperties = New ParagraphProperties(New NumberingProperties(New ParagraphStyleId With {.Val = "ListParagraph"}, New NumberingLevelReference With {.Val = item.Item1}, New NumberingId With {.Val = nid}))}
            Next

        ElseIf md.IsCodeBlock Then
            Dim mdc = CType(md, MarkdownParagraph.CodeBlock), code = mdc.Item1, lang = mdc.Item2, ignoredAfterLang = mdc.Item3
            Dim lines = code.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).ToList()
            If String.IsNullOrWhiteSpace(lines.Last) Then lines.RemoveAt(lines.Count - 1)
            If lang = "antlr" Then
                Dim run As Run = Nothing
                For Each line In lines
                    If re_grammar.IsMatch(line) Then
                        If run IsNot Nothing Then Yield New Paragraph(run) With {.ParagraphProperties = New ParagraphProperties(New ParagraphStyleId With {.Val = "Grammar"})}
                        run = New Run(New Text(line) With {.Space = SpaceProcessingModeValues.Preserve})
                    Else
                        If run Is Nothing Then run = New Run()
                        run.Append(New Break)
                        run.Append(New Text(line) With {.Space = SpaceProcessingModeValues.Preserve})
                    End If
                Next
                If run IsNot Nothing Then Yield New Paragraph(run) With {.ParagraphProperties = New ParagraphProperties(New ParagraphStyleId With {.Val = "Grammar"})}
                Return
            Else
                Dim run As New Run, onFirstLine = True
                For Each line In lines
                    If onFirstLine Then onFirstLine = False Else run.AppendChild(New Break)
                    run.AppendChild(New Text(line) With {.Space = SpaceProcessingModeValues.Preserve})
                Next
                Dim style As New ParagraphStyleId With {.Val = "Code"}
                Yield New Paragraph(run) With {.ParagraphProperties = New ParagraphProperties(style)}
                Return
            End If

        ElseIf md.IsQuotedBlock Then
            Dim mdq = CType(md, MarkdownParagraph.QuotedBlock), quoteds = mdq.Item
            Dim kind = ""
            For Each quoted In quoteds
                If quoted.IsParagraph Then
                    Dim p = CType(quoted, MarkdownParagraph.Paragraph), spans = p.Item
                    If spans.FirstOrDefault?.IsStrong Then
                        Dim strong = CType(spans.First, MarkdownSpan.Strong).Item
                        If strong.FirstOrDefault?.IsLiteral Then
                            Dim literal = CType(strong.First, MarkdownSpan.Literal).Item
                            If literal = "Annotation" Then
                                kind = "Annotation"
                                Yield New Paragraph(Span2Elements(spans.Head)) With {.ParagraphProperties = New ParagraphProperties(New ParagraphStyleId With {.Val = kind})}
                                quoted = MarkdownParagraph.NewParagraph(spans.Tail)
                            ElseIf literal = "Note" Then
                                kind = "AlertText"
                            End If
                        End If
                    End If
                    '
                    For Each qp In Paragraph2Paragraphs(quoted)
                        Dim qpp = TryCast(qp, Paragraph)
                        If qpp IsNot Nothing Then
                            Dim props As New ParagraphProperties(New ParagraphStyleId() With {.Val = kind})
                            qpp.ParagraphProperties = props
                        End If
                        Yield qp
                    Next

                ElseIf quoted.IsCodeBlock Then
                    Dim mdc = CType(quoted, MarkdownParagraph.CodeBlock), code = mdc.Item1, lang = mdc.Item2, ignoredAfterLang = mdc.Item3
                    Dim lines = code.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).ToList()
                    If String.IsNullOrWhiteSpace(lines.Last) Then lines.RemoveAt(lines.Count - 1)
                    Dim run As New Run() With {.RunProperties = New RunProperties(New RunStyle With {.Val = "CodeEmbedded"})}
                    For Each line In lines
                        If run.ChildElements.Count > 1 Then run.Append(New Break)
                        run.Append(New Text("    " & line) With {.Space = SpaceProcessingModeValues.Preserve})
                    Next
                    Yield New Paragraph(run) With {.ParagraphProperties = New ParagraphProperties(New ParagraphStyleId With {.Val = kind})}
                End If
                '
            Next
            Return

        ElseIf md.IsTableBlock Then
            Dim mdt = CType(md, MarkdownParagraph.TableBlock), header = mdt.Item1.Option, align = mdt.Item2, rows = mdt.Item3
            Dim table As New Table()
            Dim tstyle As New TableStyle With {.Val = "TableGrid"}
            Dim borders As New TableBorders()
            borders.TopBorder = New TopBorder With {.Val = BorderValues.Single}
            borders.BottomBorder = New BottomBorder With {.Val = BorderValues.Single}
            borders.LeftBorder = New LeftBorder With {.Val = BorderValues.Single}
            borders.RightBorder = New RightBorder With {.Val = BorderValues.Single}
            borders.InsideHorizontalBorder = New InsideHorizontalBorder With {.Val = BorderValues.Single}
            borders.InsideVerticalBorder = New InsideVerticalBorder With {.Val = BorderValues.Single}
            table.Append(New TableProperties(tstyle, borders))
            Dim ncols = align.Length
            For irow = -1 To rows.Length - 1
                If irow = -1 And header Is Nothing Then Continue For
                Dim mdrow = If(irow = -1, header, rows(irow))
                Dim row As New TableRow
                For icol = 0 To System.Math.Min(ncols, mdrow.Length) - 1
                    Dim mdcell = mdrow(icol)
                    Dim cell As New TableCell
                    Dim pars = Paragraphs2Paragraphs(mdcell).ToList()
                    For ip = 0 To pars.Count - 1
                        Dim p = TryCast(pars(ip), Paragraph)
                        If p IsNot Nothing Then
                            Dim props As New ParagraphProperties
                            If align(icol).IsAlignCenter Then props.Append(New Justification With {.Val = JustificationValues.Center})
                            If align(icol).IsAlignRight Then props.Append(New Justification With {.Val = JustificationValues.Right}) Dim spacing As SpacingBetweenLines = Nothing
                            If ip = pars.Count - 1 Then props.Append(New SpacingBetweenLines With {.After = "0"})
                            If props.HasChildren Then p.InsertAt(props, 0)
                        End If
                        cell.Append(pars(ip))
                    Next
                    row.Append(cell)
                Next
                table.Append(row)
            Next
            Yield table
            Return

        Else
            Yield New Paragraph(New Run(New Text($"[{md.GetType.Name}]")))
            Return
        End If
    End Function

    Function FlattenList(md As MarkdownParagraph.ListBlock) As IEnumerable(Of Tuple(Of Integer, Boolean, MarkdownParagraph.Span))
        Dim flat = FlattenList(md, 0).ToList()
        Dim isOrdered As New Dictionary(Of Integer, Boolean)
        For Each item In flat
            Dim level = item.Item1, isItemOrdered = item.Item2
            If isOrdered.ContainsKey(level) AndAlso isOrdered(level) <> isItemOrdered Then Throw New Exception("List can't mix ordered and unordered items at same level")
            If Not isOrdered.ContainsKey(level) Then isOrdered(level) = isItemOrdered
            If level > 2 Then Throw New Exception("Can't have more than 3 levels in a list")
        Next
        Return flat
    End Function

    Iterator Function FlattenList(md As MarkdownParagraph.ListBlock, level As Integer) As IEnumerable(Of Tuple(Of Integer, Boolean, MarkdownParagraph.Span))
        Dim isOrdered = md.Item1.IsOrdered, items = md.Item2
        For Each item In items
            Dim pars = item.ToList()
            Dim itemSpan As MarkdownParagraph.Span = Nothing
            Dim itemList As MarkdownParagraph.ListBlock = Nothing
            If pars.Count = 1 AndAlso pars(0).IsSpan Then
                itemSpan = CType(pars(0), MarkdownParagraph.Span)
            ElseIf pars.Count = 1 AndAlso pars(0).IsListBlock Then
                itemList = CType(pars(0), MarkdownParagraph.ListBlock)
            ElseIf pars.Count = 2 AndAlso pars(0).IsSpan AndAlso pars(1).IsListBlock Then
                itemSpan = CType(pars(0), MarkdownParagraph.Span)
                itemList = CType(pars(1), MarkdownParagraph.ListBlock)
            Else
                Throw New Exception("Nothing fancy allowed in lists")
            End If
            Yield Tuple.Create(level, isOrdered, itemSpan)
            If itemList IsNot Nothing Then
                For Each subitem In FlattenList(itemList, level + 1)
                    Yield subitem
                Next
            End If
        Next
    End Function


    Iterator Function Spans2Elements(mds As IEnumerable(Of MarkdownSpan)) As IEnumerable(Of OpenXmlElement)
        For Each md In mds
            For Each e In Span2Elements(md)
                Yield e
            Next
        Next
    End Function

    Iterator Function Span2Elements(md As MarkdownSpan) As IEnumerable(Of OpenXmlElement)
        If md.IsLiteral Then
            Dim mdl = CType(md, MarkdownSpan.Literal), s = mdl.Item

            ' workaround a bug where FSharp.Markdown doesn't parse indented codeblocks inside quotedblock
            Dim lines = s.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).ToList()
            Dim wasCode = False, run As Run = Nothing
            For Each line In lines
                If line = "" Then line = If(wasCode, "    ", " ")
                Dim isCode = line.StartsWith("    ")
                If isCode Then
                    If Not wasCode Then Yield run : run = Nothing
                    If run Is Nothing Then run = New Run With {.RunProperties = New RunProperties(New RunStyle With {.Val = "CodeEmbedded"})}
                    If Not wasCode Then run.Append(New Break, New Break)
                    run.Append(New Text(line) With {.Space = SpaceProcessingModeValues.Preserve}, New Break)
                Else
                    If wasCode Then run.Append(New Break) : Yield run : run = Nothing
                    If run Is Nothing Then run = New Run
                    run.Append(New Text(line) With {.Space = SpaceProcessingModeValues.Preserve})
                End If
                wasCode = isCode
            Next
            If run IsNot Nothing Then Yield run
            Return

        ElseIf md.IsStrong OrElse md.IsEmphasis Then
            Dim spans = If(md.IsStrong, CType(md, MarkdownSpan.Strong).Item, CType(md, MarkdownSpan.Emphasis).Item)
            For Each e In Spans2Elements(spans)
                Dim style = If(md.IsStrong, CType(New Bold, OpenXmlElement), New Italic)
                Dim run = CType(e, Run)
                If run IsNot Nothing Then run.InsertAt(New RunProperties(Style), 0)
                Yield e
            Next
            Return

        ElseIf md.IsInlineCode Then
            Dim mdi = CType(md, MarkdownSpan.InlineCode), code = mdi.Item

            Dim txt As New Text(code) With {.Space = SpaceProcessingModeValues.Preserve}
            Dim props As New RunProperties(New RunStyle With {.Val = "CodeEmbedded"})
            Dim run As New Run(txt) With {.RunProperties = props}
            Yield run
            Return

        ElseIf md.IsDirectLink Or md.IsIndirectLink Then
            Dim spans As IEnumerable(Of MarkdownSpan), url = "", alt = ""
            If md.IsDirectLink Then
                Dim mddl = CType(md, MarkdownSpan.DirectLink)
                spans = mddl.Item1
                url = mddl.Item2.Item1
                alt = mddl.Item2.Item2.Option()
            Else
                Dim mdil = CType(md, MarkdownSpan.IndirectLink), original = mdil.Item2, id = mdil.Item3
                spans = mdil.Item1
                If mddoc.DefinedLinks.ContainsKey(id) Then
                    url = mddoc.DefinedLinks(id).Item1
                    alt = mddoc.DefinedLinks(id).Item2.Option()
                End If
            End If
            Dim style = New RunStyle With {.Val = "Hyperlink"}
            Dim hyperlink As New Hyperlink With {.DocLocation = url, .Tooltip = alt}
            For Each element In Spans2Elements(spans)
                Dim run = TryCast(element, Run)
                If run IsNot Nothing Then run.InsertAt(New RunProperties(style), 0)
                hyperlink.AppendChild(run)
            Next
            Yield hyperlink
            Return

        Else
            Yield New Run(New Text($"[{md.GetType.Name}]"))
            Return
        End If
    End Function


End Class
