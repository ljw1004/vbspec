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

    ' These are the within-document links that have to be fixed up
    Dim within_document_links$ = "
A regular method is one with neither Async nor Iterator modifiers. It may be a subroutine or a function. Section 10.1.1 details what happens when a regular method is invoked.

An iterator method is one with the Iterator modifier and no Async modifier. It must be a function, and its return type must be IEnumerator, IEnumerable, or IEnumerator(Of T) or IEnumerable(Of T) for some T, and it must have no ByRef parameters. Section 10.1.2 details what happens when an iterator method is invoked.

An async method is one with the Async modifier and no Iterator modifier. It must be either a subroutine, or a function with return type Task or Task(Of T) for some T, and must have no ByRef parameters. Section 10.1.3 details what happens when an async method is invoked.

For example, when evaluating an addition operator (Section 11.3), first the left operand is evaluated, then the right operand, and then the operator itself. Blocks are executed (Section 10.1.2) by first executing their first substatement, and then proceeding one by one through the statements of the block.

4.	The method instance's control point is then set at the first statement of the method body, and the method body immediately starts to execute from there (Section 10.1.4).

Iterator methods are used as a convenient way to generate a sequence, one which can be consumed by the For Each statement. Iterator methods use the Yield statement (Section 10.15) to provide elements of the sequence. (An iterator method with no Yield statements will produce an empty 

5.	The instance's control point is then set at the first statement of the async method body, and immediately starts to execute the method body from there Section 10.1.2.

As detailed in Section 11.25 'Await Operator', execution of an Await expression has the ability to suspend the method instance's control point leaving control flow to go elsewhere

Sections 10.1.1-10.1.3 detail how and when method instances are created, and with them the copies of a method's local variables and parameters.. In addition, each time the body of a loop is entered, a new copy is made of each local variable declared inside that loop as described in Section 10.9, and the method instance now contains this copy of its local variable rather than the previous copy.

An If...Then...Else statement is the basic conditional statement. Each expression in an If...Then...Else statement must be a Boolean expression, as per Section 11.19.

A While or Do loop statement loops based on a Boolean expression. A While loop statement loops as long as the Boolean expression evaluates to true; a Do loop statement may contain a more complex condition. An expression may be placed after the Do keyword or after the Loop keyword, but not after both. The Boolean expression is evaluated as per Section 11.19. 

2.	If the loop control variable is an identifier without an As clause, then the identifier is first resolved using the simple name resolution rules (see Section 11.4.4), excepting that this occurrence of the identifier would not in and of itself cause an implicit local variable to be created (Section 10.2.1).

2.2.2.	if local variable type inference is not being used but implicit local declaration is, then an implicit local variable is created whose scope is the entire method (Section 10.2.1), and the loop control variable refers to this pre-existing variable;

Note that a new copy of the loop control variable is not created on each iteration of the loop block. In this respect, the For statement differs from For Each (Section 10.9.3).

A Catch clause with a When clause will only catch exceptions when the expression evaluates to True; the type of the expression must be a Boolean expression as per Section 11.19. 

2.	An Exit statement transfers execution to the next statement after the end of the immediately containing block statement of the specified kind. If the block is the method block, then control flow exits the method as described in Sections 10.1, 10.2, 10.3. If the Exit statement is not contained within the kind of block specified in the statement, a compile-time error occurs

Yield statements are related to iterator methods, which are described in Section 10.1.2.

The yield statement takes a single expression which must be classified as a value and whose type is implicitly convertible to the type of the iterator current variable (Section 10.1.2) of its enclosing iterator method.

An instance expression is the keyword Me. It may only be used within the body of a non-shared method, constructor, or property accessor. It is classified as a value. . The keyword Me represents the instance of the type containing the method or property accessor being executed. If a constructor explicitly invokes another constructor (Section 9.3), Me cannot be used until after that constructor call, because the instance has not yet been constructed.

For an early-bound invocation expression, the arguments are evaluated in the order in which the corresponding parameters are declared in the target method. For a late-bound member access expression, they are evaluated in the order in which they appear in the member access expression: see Section 11.3, Late-Bound Expressions.

7.	Next, if, given any two members of the set, M and N, M is more specific (Section 11.8.1.1) than N given the argument list, eliminate N from the set. If more than one member remains in the set and the remaining members are not equally specific given the argument list, a compile-time error results.

7.4.	Before type arguments have been substituted, if M is less generic (Section 11.8.1.2) than N, eliminate N from the set.

7.6.	If M and N are extension methods and M was found before N (Section 11.6.3), eliminate N from the set. For example:

7.10.	Before type arguments have been substituted, if M has greater depth of genericity (Section 11.8.1.3) than N, then eliminate N from the set. 

An array literal denotes an array whose element type, rank, and bounds are inferred from a combination of the expression context and a collection initializer. This is explained in Section 11.1.1, Expression Reclassification. For example:

The await operator is related to async methods, which are described in Section 10.1.3.

3.2.	The control point of the current async method instance is suspended, and control flow resumes in the current caller (defined in Section 10.1.3).

3.3.2.	then it resumes control flow at the control point of the async method instance (see Section 10.1.3),
"

    Sub Main()
        Dim fn = PickUniqueFilename("webmd.docx")
        Dim files = {"introduction.md", "lexical-grammar.md", "preprocessing-directives.md",
                     "general-concepts.md", "attributes.md", "source-files-And-namespaces.md", "types.md",
                     "conversions.md", "type-members.md", "statements.md", "expressions.md",
                     "documentation-comments.md"}
        Dim md = MarkdownSpec.ReadFiles(From file In files Select $"..\..\..\vb\{file}")

        Dim grammar = Antlr.ReadFile("vb11.g4")
        If Not grammar.IsSameAs(md.Grammar) Then Throw New Exception("Grammar mismatch")
        Html.WriteFile(md.Grammar, "vb11", "vb11.html")

        md.WriteFile("vb-template.docx", fn)
        Process.Start("vb11.html")
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

