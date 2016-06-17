' This "driver" project aims to provide a within-VS IDE for working on the spec.
'
' As a pre-build step it regenerates the spec docx and opens it up in Word.
' What's useful here is that errors are reported in the VS Error List,
' and you can double-click them to jump to the relevant line in the .md file.
'    mdspec2docx.exe *.md *.g4 template.docx -o vb.html -o "Visual Basic Language Specification.docx" -td vs\obj
'
' Then you can run this console app to do the final packaging:
'    * launches Word to update the TOC with page numbers
'    * launch esWord to export it as PDF


Imports System.IO
Imports Microsoft.Office.Interop.Word

Module Module1
    Function Main() As Integer
        Dim fn = "Visual Basic Language Specification.docx"
        If Not File.Exists(fn) Then Directory.SetCurrentDirectory("..\..\..")
        If Not File.Exists(fn) Then Console.Error.WriteLine($"Can't find '{fn}'") : Return 1

        fn = Path.GetFullPath(fn)
        Try
            Console.WriteLine("Launching Word")
            Dim word As New Application
            Try
                Console.WriteLine($"Opening '{Path.GetFileName(fn)}'")
                Dim doc = word.Documents.Open(CObj(fn), AddToRecentFiles:=False, Visible:=True)
                doc.Activate()
                Try
                    If doc.TablesOfContents.Count >= 1 Then
                        Console.WriteLine("Updating TOC with page numbers")
                        Dim toc = doc.TablesOfContents(1)
                        toc.RightAlignPageNumbers = True
                        toc.UseHeadingStyles = True
                        toc.UpperHeadingLevel = 1
                        toc.LowerHeadingLevel = 2
                        toc.IncludePageNumbers = True
                        toc.UseHyperlinks = True
                        toc.TabLeader = WdTabLeader.wdTabLeaderDots
                        toc.Update()
                        toc.UpdatePageNumbers()
                        Console.WriteLine($"Saving '{Path.GetFileName(fn)}'")
                        doc.Save()
                        Dim pdf = Path.ChangeExtension(fn, ".pdf")
                        Console.WriteLine($"Saving '{Path.GetFileName(pdf)}'")
                        doc.SaveAs2(CObj(pdf), WdSaveFormat.wdFormatPDF)
                    End If
                Finally
                    Console.WriteLine("Closing")
                    doc.Close()
                End Try
            Finally
                word.Quit()
            End Try
            Console.WriteLine("Done.")
            Return 0
        Catch ex As Exception
            Console.WriteLine($"ERROR - {ex.Message}")
            Return 1
        End Try
    End Function

End Module
