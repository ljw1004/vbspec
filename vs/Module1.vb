' This "driver" project aims to provide a within-VS IDE for working on the spec.
'
' Run this driver in DEBUG mode to regenerate the spec docx and quickly see how it looks:
'    mdspec2docx.exe *.md *.g4 template.docx -o vb.html -o "Visual Basic Language Specification.docx"
'
' Run it in RELEASE mode to automate updating the docx and pdf for final publication:
'    * launches Word to update the TOC with page numbers
'    * launch esWord to export it as PDF


' TODO:
' (1) Change license preamble in readme.md and template.docx to match the rest of Roslyn,
'     and change license.txt to match Roslyn
' (2) Error if the readme.md headings fail to contain every level1+level2 heading
'     and consider adding OverloadResolution + TypeInference
' (3) obj directory
' (4) Prepare PR for Roslyn

' Rebuild this project to run the "mdspec2docx" tool.
' Any errors will be shown in the Error List.
' If no errors, then it will pop up the new vb.html and language specification.docx files.

Imports System.IO
Imports Microsoft.Office.Interop.Word

Module Module1
    Function Main() As Integer
        Return MainAsync().GetAwaiter().GetResult()
    End Function

    Async Function MainAsync() As Task(Of Integer)
        If Not File.Exists("readme.md") Then Directory.SetCurrentDirectory("..\..\..")
        If Not File.Exists("readme.md") Then Console.Error.WriteLine("Can't find 'readme.md'") : Return 1

        Dim md2docx = "vs\packages\mdspec2docx.1.0.0\tools\mdspec2docx.exe"
        If Not File.Exists(md2docx) Then Console.Error.WriteLine($"Can't find '{Path.GetFileName(md2docx)}'") : Return 1

#If DEBUG Then
        Console.WriteLine("DEBUG build - will only preview the docx")
        Console.WriteLine("(if you want to update TOC and generate PDF, switch to RELEASE)")
#End If

        Dim p As New Process
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.FileName = md2docx
        p.StartInfo.Arguments = "*.md *.g4 template.docx -o vb.html -o ""Visual Basic Language Specification.docx"""
        p.Start()
        Dim task1 = CopyStreamAsync(p.StandardError, Console.Error)
        Dim task2 = CopyStreamAsync(p.StandardOutput, Console.Out)
        Await Threading.Tasks.Task.WhenAll(task1, task2)
        p.WaitForExit()
        If p.ExitCode <> 0 Then Return p.ExitCode
        Dim fn = "Visual Basic Language Specification.docx"

#If DEBUG Then
        Process.Start(fn)
        Return 0
#End If

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

    Async Function CopyStreamAsync(read As StreamReader, write As TextWriter) As Threading.Tasks.Task
        While True
            Dim s = Await read.ReadLineAsync()
            If s Is Nothing Then Return
            Await write.WriteLineAsync(s)
        End While
    End Function

End Module
