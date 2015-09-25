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
        Using templateDoc = WordprocessingDocument.Open("vb-template.docx", False),
                resultDoc = WordprocessingDocument.Create("vbmd.docx", WordprocessingDocumentType.Document)
            For Each part In templateDoc.Parts
                resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId)
            Next

            Dim body = resultDoc.MainDocumentPart.Document.Body
            body.RemoveAllChildren()

            'Dim src = "|------------|-----------|-------|-----|
            '| AddHandler | AddressOf | Alias | And | 
            '| AndAlso | As | Boolean | ByRef | 
            '"
            'For Each p In New MarkdownParser(src).Paragraphs
            '    body.AppendChild(p)
            'Next

            For Each p In New MarkdownParser(File.ReadAllText("..\..\..\vb\introduction.md")).Paragraphs
                body.AppendChild(p)
            Next
            For Each p In New MarkdownParser(File.ReadAllText("..\..\..\vb\lexical-grammar.md")).Paragraphs
                body.AppendChild(p)
            Next
        End Using

        Process.Start("vbmd.docx")
    End Sub

    <Extension>
    Public Function [Option](Of T As Class)(this As FSharpOption(Of T)) As T
        If FSharpOption(Of T).GetTag(this) = FSharpOption(Of T).Tags.None Then Return Nothing
        Return this.Value
    End Function

End Module

Class MarkdownParser
    Dim doc As MarkdownDocument

    Sub New(s As String)
        doc = Markdown.Parse(s)
    End Sub

    Function Paragraphs() As IEnumerable(Of OpenXmlCompositeElement)
        Return Paragraphs2Paragraphs(doc.Paragraphs)
    End Function

    Iterator Function Paragraphs2Paragraphs(pars As IEnumerable(Of MarkdownParagraph)) As IEnumerable(Of OpenXmlCompositeElement)
        For Each md In pars
            For Each p In Paragraph2Paragraphs(md)
                Yield p
            Next
        Next
    End Function


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

        ElseIf md.IsCodeBlock Then
            Dim mdc = CType(md, MarkdownParagraph.CodeBlock), code = mdc.Item1, lang = mdc.Item2, ignoredAfterLang = mdc.Item3
            Dim lines = code.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).ToList()
            If String.IsNullOrWhiteSpace(lines.Last) Then lines.RemoveAt(lines.Count - 1)
            Dim grammar = IsGrammar(lines)
            If grammar Then
                For Each line In lines
                    Dim run As New Run(New Text(line))
                    Dim style As New ParagraphStyleId With {.Val = "Grammar"}
                    Yield New Paragraph(run) With {.ParagraphProperties = New ParagraphProperties(style)}
                Next
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
            Dim mdq = CType(md, MarkdownParagraph.QuotedBlock), pars = mdq.Item
            If Not pars.Any Then Return
            Dim kind = ""
            If pars.Count = 1 AndAlso pars.First.IsParagraph Then
                Dim firstPar = CType(pars.First, MarkdownParagraph.Paragraph), spans = firstPar.Item
                If spans.Any AndAlso spans.First.IsStrong Then
                    Dim strong = CType(spans.First, MarkdownSpan.Strong).Item
                    If strong.Any AndAlso strong.First.IsLiteral Then
                        Dim literal = CType(strong.First, MarkdownSpan.Literal).Item
                        If literal = "Note" Then kind = "AlertText"
                        If literal = "Annotation" Then kind = "Annotation"
                    End If
                End If
            End If

            If kind = "Annotation" Then
                Dim props As New ParagraphProperties(New ParagraphStyleId With {.Val = kind})
                Dim onlyPar = CType(pars.First, MarkdownParagraph.Paragraph)
                Dim strong = onlyPar.Item.First, rest = onlyPar.Item.Skip(1).ToList()
                Dim literal = TryCast(rest(0), MarkdownSpan.Literal)
                rest(0) = MarkdownSpan.NewLiteral(literal.Item.TrimStart())

                Dim props1 As New ParagraphProperties(New ParagraphStyleId With {.Val = kind})
                Dim p1 As New Paragraph(Span2Elements(strong)) With {.ParagraphProperties = props1}
                Yield p1

                Dim props2 As New ParagraphProperties(New ParagraphStyleId With {.Val = kind})
                Dim p2 As New Paragraph(Spans2Elements(rest)) With {.ParagraphProperties = props2}
                Yield p2
                Return
            Else
                For Each p In Paragraphs2Paragraphs(pars)
                    Dim pp = TryCast(p, Paragraph)
                    If pp IsNot Nothing Then
                        Dim props As New ParagraphProperties(New ParagraphStyleId() With {.Val = kind})
                        pp.ParagraphProperties = props
                    End If
                    Yield p
                Next
                Return
            End If

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



    Function IsGrammar(lines As IEnumerable(Of String)) As Boolean
        Static Dim re As New Regex("^([a-zA-Z]+): ")
        If Not lines.Any Then Return False
        Dim s = lines.First
        If re.IsMatch(s) Then Return True
        Return False
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
            Dim txt As New Text(s) With {.Space = SpaceProcessingModeValues.Preserve}
            Yield New Run(txt)
            Return
        ElseIf md.IsStrong OrElse md.IsEmphasis Then
            Dim spans = If(md.IsStrong, CType(md, MarkdownSpan.Strong).Item, CType(md, MarkdownSpan.Emphasis).Item)
            Dim style = If(md.IsStrong, CType(New Bold, OpenXmlElement), New Italic)
            For Each e In Spans2Elements(spans)
                Dim run = CType(e, Run)
                If run IsNot Nothing Then run.InsertAt(New RunProperties(style), 0)
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
                If doc.DefinedLinks.ContainsKey(id) Then
                    url = doc.DefinedLinks(id).Item1
                    alt = doc.DefinedLinks(id).Item2.Option()
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
