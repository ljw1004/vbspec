Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports FSharp.Markdown
Imports Microsoft.FSharp.Core

' TODO: clean up within-document links
' TODO: restore frontmatter
' TODO: table of contents
' TODO: at end of statements.md, why does Yield grammar have a leading blank line? and Await at end of expressions.md?
' TODO: missing space when grammar follows a table
' TODO: spec references in html grammar

Module Module1

    Sub Main()
        Dim grammar = Antlr.ReadFile("vb11.g4")
        Html.WriteFile(grammar, "vb11", "vb11.html")

        Dim fn = PickUniqueFilename("webmd.docx")
        Dim files = {"introduction.md", "lexical-grammar.md", "preprocessing-directives.md",
                     "general-concepts.md", "attributes.md", "source-files-And-namespaces.md", "types.md",
                     "conversions.md", "type-members.md", "statements.md", "expressions.md",
                     "documentation-comments.md"}
        Dim md = MarkdownSpec.ReadFiles(From file In files Select $"..\..\..\vb\{file}")
        If Not grammar.IsSameAs(Antlr.ReadString(md.AntlrCodeBlocks)) Then Throw New Exception("Grammar mismatch")
        md.WriteFile("vb-template.docx", fn)
        Process.Start(fn)
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
        Loop
    End Function

End Module

