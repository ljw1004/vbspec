VB spec in markdown
=====================

Currently the authoritative VB spec is written in Word in English. The VB language design team edit it for each release cycle. As a "Track Changes" Word document it gets sent to translators, who maintain translations in five or so languages. It then gets installed as part of Visual Studio in `C:\Program Files (x86)\Microsoft Visual Studio 14.0\VB\Specifications\1033\Visual Basic Language Specification.docx`. (Actually, in VS2015, we failed to do this and the version currently there dates back to VS2012).

* I want to move the authoritative version over to Markdown, to be stored on the Roslyn github.
* I want it to be easy for folks to hyperlink to sections of the spec
* I want it to accept pull-requests
* I want some way to track implementor-notes, for instance "Here's a subtle piece of code that illustrates this paragraph".
* I want to keep the existing translator workflow based on "Track Changes" Word documents
* I want to keep a single-document version available for download
* I want to maintain fewer copies of the grammar, and I want the grammar to be machine-navigable, and easier for humans to navigate.

So here's my step-by-step plan...

1. [Scott Dorman](https://github.com/scottdorman) kindly made the first pass at converting the Word document into a series of markdown files.
2. I'm in the process of writing a converter which parses the markdown generates a Word document that's closely matched in style to the existing VB spec. And in the process of editing the markdown files to use consistent conventions.
3. I will maintain a standalone file called `vb-grammar.g` which is in ANTLR3 format. I'll copy+paste sections of this into the markdown spec, adopting ANTLR as the grammar notation for the VB spec. I'll also resurrect my hyperlinked VB grammar document, which is produced from the ANTLR spec.
4. The Word document will not have a single grammar file at the end (it's redundant and better served online), nor a "change history section at the end (this is better done through github change tracking)
5. Once everything's done, and is close enough to the current VB spec, then I'll check it into the Roslyn github.
6. I'll take a Word snapshot of the VB spec as of VS2012, which is what the spec currently sits at.
7. I'll author all the changes that came with VS2015 directly into github, in markdown.
8. I'll take another Word snapshot of the VB spec as of VS2015, compare it to the previous snapshot, and synthesize a Word "Track Changes" document out of this to send to the translator folks.
9. I'll find a good MSDN location to host Word versions of the VS2015 spec, and have it also link to the Roslyn github markdown version of the spec.



# Conventions

Notes are like this:
```
> __Note__
> Here is the body of the note
```

Annotations are like this:
```
> __Annotation__
> Here is the body of the annotation
```

VB code blocks are like this:
```
``vb
Dim x = "Hello"
``
```

To put a code block inside an annotation, it needs a quoted but otherwise empty line before and after:
```
> __Annotation__
> Start of annotation
>
> ``vb
> Dim x = "Hello"
> ``
>
> More annotation
```

Grammar blocks are like this:
```
``antlr
Start
    : Left
    | Right
    ;

Left
    : 'hello'
    ;
```

Links are like this, which will render in github as a hyperlinked word "Conventions", and will render in Word as "5.2: Conventions". The thing that comes after the # is the section/subsection title, stripping everything other than alphanumerics and hyphen and underscore, and converting to lowercase.
```
For more information see Section [Conventions](README.md#conventions)
```

## Experiments and stuff

The following is [link1](README.md#conventions)

The following is [link2](README.md#experiments-and-stuff)

The following is a heading with inline code. Github gives it the link [link3](README.md#c)

First try a heading with a codeblock:

### `<c>`

Next try a heading with ampersand-escapes for the lt/gt characters:

### &lt;c&gt;

Next try a heading with just lt/gt characters:

### <c>

Next, that was the end of my heading experiments. Back to links.

The following is a heading with numbers. Github gives it the link [link4](README.md#123-hello-456-world)

### 123 hello 456 world

The following is a heading with symbols. Github gives it the link [link5](README.md#abcdefghijk_l-mnopqrstuvwxyz). It preserves hyphen, underscore, numerics, alphas (converted to lowercase), and removes the rest.

### a!b@c#d$e%f^g&h*i(j)k_l-m+n=o[p{q|r\s;t:u'v"w,x.y?z/

The following experiment is about code blocks in quoted blocks. Both styles are fine.

> __Annotation__
> This is an annotation
> ```vb
> Module Module1
> End Module
> ```
> More annotation

and this is another one...

> __Annotation__
> This is an annotation

> ```vb
> Module Module1
> End Module
> ```

> More annotation

and this is another one...

> __Annotation__
> This is an annotation with five spaces in the nested code block

>     This is a nested code block
>     of multiple lines

> Continued annotation

and another one...

> __Annotation__
> This is an annotation with four spaces in the nested code block

>    This is a nested code block
>    of multiple lines

> Continued annotation


-------------------------------------------------------------

The following experiment is about nested lists. It shows that github markdown ignores the number you've written (using it solely to infer "ordered" vs "unordered") and the numbering scheme it uses is
```
1. Level One
   i. Level Two
      b. Level 3
```

1.  Hello
2.  World
    1. Alpha
        1. One
        2. Two
    2. Beta
        1. OneB
        2. TwoB
3.  Goodbye

and again...

1.  Hello
2.  World
    1. Alpha
    2. Beta

    Continuation
3.  Goodbye

-------------------------------------------------------------

The following experiment shows we must use brace codeblocks, not indented codeblocks. The brace languages I'll use will be "vb" and "antlr".

Here's a small experiment...

VB inline code `Dim x As Integer = 5`

C# inline code `int x = 5;`

ANTRL inline code `Start: Left | Right | 'a'`

VB brace code "vb"
```vb
Dim x As Integer = 5
Dim y As String = "hello"
```

VB brace code "vb.net"
```vb.net
Dim x As Integer = 5
Dim y As String = "hello"
```

VB indented code

    Dim x As Integer = 5
    Dim y As String = "hello"

C# brace code "cs"
```cs
int x = 5;
string y = "hello";
```

C# brace code "csharp"
```csharp
int x = 5;
string y = "hello";
```

C# indented code

    int x = 5;
    string y = "hello";

ANTLR brace code "antlr"
```antlr
Start: Left | Right | 'hello';
Left: 'world';
Right: 'there';
```

ANTLR brace code "ANTLR"
```ANTLR
Start: Left | Right | 'hello';
Left: 'world';
Right: 'there';
```

ANTLR indented code

    Start: Left | Right | 'hello';
    Left: 'world';
    Right: 'there';

-------------------------------------------------------------

End of experiments.
