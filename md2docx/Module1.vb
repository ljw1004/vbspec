Module Module1

    Sub Main()
        Dim fn = PickUniqueFilename("webmd.docx")
        Dim files = {"introduction.md", "lexical-grammar.md", "preprocessing-directives.md",
                     "general-concepts.md", "attributes.md", "source-files-And-namespaces.md", "types.md",
                     "conversions.md", "type-members.md", "statements.md", "expressions.md",
                     "documentation-comments.md"}
        Dim md = MarkdownSpec.ReadFiles(From file In files Select $"..\..\..\vb\{file}")

        Dim grammar = Antlr.ReadFile("vb.g4")
        If Not grammar.AreProductionsSameAs(md.Grammar) Then Throw New Exception("Grammar mismatch")
        md.Grammar.Name = grammar.Name
        Html.WriteFile(md.Grammar, "vb.html")

        md.WriteFile("vb-template.docx", fn)
        Process.Start(fn)
        Process.Start("vb.html")
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

