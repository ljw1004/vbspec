Module Module1

    Sub Main()
        ' Read readme.md to find a list of files, and read them in
        Dim readme = FSharp.Markdown.Markdown.Parse(IO.File.ReadAllText("readme.md"))
        Dim files = (From list In readme.Paragraphs.OfType(Of FSharp.Markdown.MarkdownParagraph.ListBlock)
                     Let items = list.Item2
                     From par In items
                     From spanpar In par.OfType(Of FSharp.Markdown.MarkdownParagraph.Span)
                     Let spans = spanpar.Item
                     From link In spans.OfType(Of FSharp.Markdown.MarkdownSpan.DirectLink)
                     Let url = link.Item2.Item1
                     Select url).ToList.Distinct
        Dim md = MarkdownSpec.ReadFiles(files)

        ' Now md.Gramar contains the grammar as extracted out of the *.md files, and moreover has
        ' correct references to within the spec. We'll check that it has the same productions as
        ' in the corresponding ANTLR file
        Dim antlrfn = IO.Directory.GetFiles(".", "*.g4").First
        Dim htmlfn = IO.Path.ChangeExtension(antlrfn, ".html")
        Dim grammar = Antlr.ReadFile(antlrfn)
        If Not grammar.AreProductionsSameAs(md.Grammar) Then Throw New Exception("Grammar mismatch")
        md.Grammar.Name = grammar.Name

        ' Generate the Specification.docx file
        Dim fn = PickUniqueFilename("Specification.docx")
        md.WriteFile("template.docx", fn)
        Process.Start(fn)

        ' Generate the grammar.html file
        Html.WriteFile(md.Grammar, htmlfn)
        Process.Start(htmlfn)
    End Sub


    Function PickUniqueFilename(suggestion As String) As String
        Dim base = IO.Path.GetFileNameWithoutExtension(suggestion)
        Dim ext = IO.Path.GetExtension(suggestion)

        Dim ifn = 0
        Do
            Dim fn = base & If(ifn = 0, "", CStr(ifn)) & ext
            If Not IO.File.Exists(fn) Then Return fn
            Try
                IO.File.Delete(fn) : Return fn
            Catch ex As Exception
            End Try
            ifn += 1
        Loop
    End Function

End Module

