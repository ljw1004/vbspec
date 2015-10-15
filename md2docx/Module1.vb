Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports FSharp.Markdown
Imports Microsoft.FSharp.Core

' TODO: table of contents
'   http://stackoverflow.com/questions/9762684/how-to-generate-table-of-contents-using-openxml-sdk-2-0
'   http://stackoverflow.com/questions/8560753/update-toc-in-docx-document-using-documentformat-openxml-c
'   http://stackoverflow.com/questions/2208640/update-word-docs-table-of-contents-part-automatically-using-net
'   http://stackoverflow.com/questions/17486741/table-of-content-for-microsoft-interop-word-dll
' TODO: restore frontmatter

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
        Dim base = Path.GetFileNameWithoutExtension(suggestion)
        Dim ext = Path.GetExtension(suggestion)

        Dim ifn = 0
        Do
            Dim fn = base & If(ifn = 0, "", CStr(ifn)) & ext
            If Not File.Exists(fn) Then Return fn
            Try
                File.Delete(fn) : Return fn
            Catch ex As Exception
            End Try
            ifn += 1
        Loop
    End Function

End Module

