The Microsoft Visual Basic Language Specification
=================
__Version 11.0__

> Paul Vick, Lucian Wischik

> Microsoft Corporation

> The information contained in this document represents the current view of Microsoft Corporation on the issues discussed as of the date of publication. Because Microsoft must respond to changing market conditions, it should not be interpreted to be a commitment on the part of Microsoft, and Microsoft cannot guarantee the accuracy of any information presented after the date of publication.

> This Language Specification is for informational purposes only. MICROSOFT MAKES NO WARRANTIES, EXPRESS, IMPLIED OR STATUTORY, AS TO THE INFORMATION IN THIS DOCUMENT.

> Complying with all applicable copyright laws is the responsibility of the user. Without limiting the rights under copyright, no part of this document may be reproduced, stored in or introduced into a retrieval system, or transmitted in any form or by any means (electronic, mechanical, photocopying, recording, or otherwise), or for any purpose, without the express written permission of Microsoft Corporation.

> Microsoft may have patents, patent applications, trademarks, copyrights, or other intellectual property rights covering subject matter in this document. Except as expressly provided in any written license agreement from Microsoft, the furnishing of this document does not give you any license to these patents, trademarks, copyrights, or other intellectual property.

> Unless otherwise noted, the example companies, organizations, products, domain names, e-mail addresses, logos, people, places and events depicted herein are fictitious, and no association with any real company, organization, product, domain name, email address, logo, person, place or event is intended or should be inferred.

> &copy; 2012 Microsoft Corporation. All rights reserved.

> Microsoft, MS-DOS, Visual Basic, Windows 2000, Windows 95, Windows 98, Windows ME, Windows NT, Windows XP, Windows Vista and Windows are either registered trademarks or trademarks of Microsoft Corporation in the United States and/or other countries.

> The names of actual companies and products mentioned herein may be the trademarks of their respective owners.

##Table of Contents

1. [Introduction](introduction.md)<br/>
1.1 [Grammar Notation](introduction.md)<br/>
1.2 [Compatibility](introduction.md)<br/>
1.2.1 [Kinds of compatibility breaks](introduction.md)<br/>
1.2.2 [Impact Criteria](introduction.md)<br/>
1.2.3 [Language deprecation](introduction.md)<br/>
2. [Lexical Grammar](lexical-grammar.md)<br/>
2.1 [Characters and Lines](lexical-grammar.md)<br/>
2.1.1 [Line Terminators](lexical-grammar.md)<br/>
2.1.2 [Line Continuation](lexical-grammar.md)<br/>
2.1.3 [White Space](lexical-grammar.md)<br/>
2.1.4 [Comments](lexical-grammar.md)<br/>
2.2 [Identifiers](lexical-grammar.md)<br/>
2.2.1 [Type Characters](lexical-grammar.md)<br/>
2.3 [Keywords](lexical-grammar.md)<br/>
2.4 [Literals](lexical-grammar.md)<br/>
2.4.1 [Boolean Literals](lexical-grammar.md)<br/>
2.4.2 [Integer Literals](lexical-grammar.md)<br/>
2.4.3 [Floating-Point Literals](lexical-grammar.md)<br/>
2.4.4 [String Literals](lexical-grammar.md)<br/>
2.4.5 [Character Literals](lexical-grammar.md)<br/>
2.4.6 [Date Literals](lexical-grammar.md)<br/>
2.4.7 [Nothing](lexical-grammar.md)<br/>
2.5 [Separators](lexical-grammar.md)<br/>
2.6 [Operator Characters](lexical-grammar.md)<br/>
3. [Preprocessing Directives](preprocessing-directives.md)<br/>
3.1 [Conditional Compilation](preprocessing-directives.md)<br/>
3.1.1 [Conditional Constant Directives](preprocessing-directives.md)<br/>
3.1.2 [Conditional Compilation Directives](preprocessing-directives.md)<br/>
3.2 [External Source Directives](preprocessing-directives.md)<br/>
3.3 [Region Directives](preprocessing-directives.md)<br/>
3.4 [External Checksum Directives](preprocessing-directives.md)<br/>
4. [General Concepts](general-concepts.md)<br/>
4.1 [Declarations](general-concepts.md)<br/>
4.1.1 [Overloading and Signatures](general-concepts.md)<br/>
4.2 [Scope](general-concepts.md)<br/>
4.3 [Inheritance](general-concepts.md)<br/>
4.3.1 [MustInherit and NotInheritable Classes](general-concepts.md)<br/>
4.3.2 [Interfaces and Multiple Inheritance](general-concepts.md)<br/>
4.3.3 [Shadowing](general-concepts.md)<br/>
4.4 [Implementation](general-concepts.md)<br/>
4.4.1 [Implementing Methods](general-concepts.md)<br/>
4.5 [Polymorphism](general-concepts.md)<br/>
4.5.1 [Overriding Methods](general-concepts.md)<br/>
4.6 [Accessibility](general-concepts.md)<br/>
4.6.1 [Constituent Types](general-concepts.md)<br/>
4.7 [Type and Namespace Names](general-concepts.md)<br/>
4.7.1 [Qualified Name Resolution for namespaces and types](general-concepts.md)<br/>
4.7.2 [Unqualified Name Resolution for namespaces and types](general-concepts.md)<br/>
4.8 [Variables](general-concepts.md)<br/>
4.9 [Generic Types and Methods](general-concepts.md)<br/>
4.9.1 [Type Parameters](general-concepts.md)<br/>
4.9.2 [Type Constraints](general-concepts.md)<br/>
4.9.3 [Type Parameter Variance](general-concepts.md)<br/>
5. [Attributes](attributes.md)<br/>
5.1 [Attribute Classes](attributes.md)<br/>
5.2 [Attribute Blocks](attributes.md)<br/>
5.2.1 [Attribute Names](attributes.md)<br/>
5.2.2 [Attribute Arguments](attributes.md)<br/>
6. [Source Files and Namespaces](source-files-and-namespaces.md)<br/>
6.1 [Program Startup and Termination](source-files-and-namespaces.md)<br/>
6.2 [Compilation Options](source-files-and-namespaces.md)<br/>
6.2.1 [Option Explicit Statement](source-files-and-namespaces.md)<br/>
6.2.2 [Option Strict Statement](source-files-and-namespaces.md)<br/>
6.2.3 [Option Compare Statement](source-files-and-namespaces.md)<br/>
6.2.4 [Integer Overflow Checks](source-files-and-namespaces.md)<br/>
6.2.5 [Option Infer Statement](source-files-and-namespaces.md)<br/>
6.3 [Imports Statement](source-files-and-namespaces.md)<br/>
6.3.1 [Import Aliases](source-files-and-namespaces.md)<br/>
6.3.2 [Namespace Imports](source-files-and-namespaces.md)<br/>
6.3.3 [XML Namespace Imports](source-files-and-namespaces.md)<br/>
6.4 [Namespaces](source-files-and-namespaces.md)<br/>
6.4.1 [Namespace Declarations](source-files-and-namespaces.md)<br/>
6.4.2 [Namespace Members](source-files-and-namespaces.md)<br/>
7. [Types](types.md)<br/>
7.1 [Value Types and Reference Types](types.md)<br/>
7.1.1 [Nullable Value Types](types.md)<br/>
7.2 [Interface Implementation](types.md)<br/>
7.3 [Primitive Types](types.md)<br/>
7.4 [Enumerations](types.md)<br/>
7.4.1 [Enumeration Members](types.md)<br/>
7.4.2 [Enumeration Values](types.md)<br/>
7.5 [Classes](types.md)<br/>
7.5.1 [Class Base Specification](types.md)<br/>
7.5.2 [Class Members](types.md)<br/>
7.6 [Structures](types.md)<br/>
7.6.1 [Structure Members](types.md)<br/>
7.7 [Standard Modules](types.md)<br/>
7.7.1 [Standard Module Members](types.md)<br/>
7.8 [Interfaces](types.md)<br/>
7.8.1 [Interface Inheritance](types.md)<br/>
7.8.2 [Interface Members](types.md)<br/>
7.9 [Arrays](types.md)<br/>
7.10 [Delegates](types.md)<br/>
7.11 [Partial types](types.md)<br/>
7.12 [Constructed Types](types.md)<br/>
7.12.1 [Open Types and Closed Types](types.md)<br/>
7.13 [Special Types](types.md)<br/>
8. [Conversions](conversions.md)<br/>
8.1 [Implicit and Explicit Conversions](conversions.md)<br/>
8.2 [Boolean Conversions](conversions.md)<br/>
8.3 [Numeric Conversions](conversions.md)<br/>
8.4 [Reference Conversions](conversions.md)<br/>
8.4.1 [Reference Variance Conversions](conversions.md)<br/>
8.4.2 [Anonymous Delegate Conversions](conversions.md)<br/>
8.5 [Array Conversions](conversions.md)<br/>
8.6 [Value Type Conversions](conversions.md)<br/>
8.6.1 [Nullable Value Type Conversions](conversions.md)<br/>
8.7 [String Conversions](conversions.md)<br/>
8.8 [Widening Conversions](conversions.md)<br/>
8.9 [Narrowing Conversions](conversions.md)<br/>
8.10 [Type Parameter Conversions](conversions.md)<br/>
8.11 [User-Defined Conversions](conversions.md)<br/>
8.11.1 [Most Specific Widening Conversion](conversions.md)<br/>
8.11.2 [Most Specific Narrowing Conversion](conversions.md)<br/>
8.12 [Native Conversions](conversions.md)<br/>
8.13 [Dominant Type](conversions.md)<br/>
9. [Type Members](type-members.md)<br/>
9.1 [Interface Method Implementation](type-members.md)<br/>
9.2 [Methods](type-members.md)<br/>
9.2.1 [Regular, Async and Iterator Method Declarations](type-members.md)<br/>
9.2.2 [External Method Declarations](type-members.md)<br/>
9.2.3 [Overridable Methods](type-members.md)<br/>
9.2.4 [Shared Methods](type-members.md)<br/>
9.2.5 [Method Parameters](type-members.md)<br/>
9.2.5.1 [Value Parameters](type-members.md)<br/>
9.2.5.2 [Reference Parameters](type-members.md)<br/>
9.2.5.3 [Optional Parameters](type-members.md)<br/>
9.2.5.4 [ParamArray Parameters](type-members.md)<br/>
9.2.6 [Event Handling](type-members.md)<br/>
9.2.7 [Extension Methods](type-members.md)<br/>
9.2.8 [Partial Methods](type-members.md)<br/>
9.3 [Constructors](type-members.md)<br/>
9.3.1 [Instance Constructors](type-members.md)<br/>
9.3.2 [Shared Constructors](type-members.md)<br/>
9.4 [Events](type-members.md)<br/>
9.4.1 [Custom Events](type-members.md)<br/>
9.5 [Constants](type-members.md)<br/>
9.6 [Instance and Shared Variables](type-members.md)<br/>
9.6.1 [Read-Only Variables](type-members.md)<br/>
9.6.2 [WithEvents Variables](type-members.md)<br/>
9.6.3 [Variable Initializers](type-members.md)<br/>
9.6.3.1 [Regular Initializers](type-members.md)<br/>
9.6.3.2 [Object Initializers](type-members.md)<br/>
9.6.3.3 [Array-Size Initializers](type-members.md)<br/>
9.6.4 [System.MarshalByRefObject Classes](type-members.md)<br/>
9.7 [Properties](type-members.md)<br/>
9.7.1 [Get Accessor Declarations](type-members.md)<br/>
9.7.2 [Set Accessor Declarations](type-members.md)<br/>
9.7.3 [Default Properties](type-members.md)<br/>
9.7.4 [Automatically Implemented Properties](type-members.md)<br/>
9.7.5 [Iterator Properties](type-members.md)<br/>
9.8 [Operators](type-members.md)<br/>
9.8.1 [Unary Operators](type-members.md)<br/>
9.8.2 [Binary Operators](type-members.md)<br/>
9.8.3 [Conversion Operators](type-members.md)<br/>
9.8.4 [Operator Mapping](type-members.md)<br/>
10. [Statements](statements.md)<br/>
10.1 [Control Flow](statements.md)<br/>
10.1.1 [Regular Methods](statements.md)<br/>
10.1.2 [Iterator Methods](statements.md)<br/>
10.1.3 [Async methods](statements.md)<br/>
10.1.4 [Blocks and Labels](statements.md)<br/>
10.1.5 [Local Variables and Parameters](statements.md)<br/>
10.2 [Local Declaration Statements](statements.md)<br/>
10.2.1 [Implicit Local Declarations](statements.md)<br/>
10.3 [With Statement](statements.md)<br/>
10.4 [SyncLock Statement](statements.md)<br/>
10.5 [Event Statements](statements.md)<br/>
10.5.1 [RaiseEvent Statement](statements.md)<br/>
10.5.2 [AddHandler and RemoveHandler Statements](statements.md)<br/>
10.6 [Assignment Statements](statements.md)<br/>
10.6.1 [Regular Assignment Statements](statements.md)<br/>
10.6.2 [Compound Assignment Statements](statements.md)<br/>
10.6.3 [Mid Assignment Statement](statements.md)<br/>
10.7 [Invocation Statements](statements.md)<br/>
10.8 [Conditional Statements](statements.md)<br/>
10.8.1 [If...Then...Else Statements](statements.md)<br/>
10.8.2 [Select Case Statements](statements.md)<br/>
10.9 [Loop Statements](statements.md)<br/>
10.9.1 [While...End While and Do...Loop Statements](statements.md)<br/>
10.9.2 [For...Next Statements](statements.md)<br/>
10.9.3 [For Each...Next Statements](statements.md)<br/>
10.10 [Exception-Handling Statements](statements.md)<br/>
10.10.1 [Structured Exception-Handling Statements](statements.md)<br/>
10.10.1.1 [Finally Blocks](statements.md)<br/>
10.10.1.2 [Catch Blocks](statements.md)<br/>
10.10.1.3 [Throw Statement](statements.md)<br/>
10.10.2 [Unstructured Exception-Handling Statements](statements.md)<br/>
10.10.2.1 [Error Statement](statements.md)<br/>
10.10.2.2 [On Error Statement](statements.md)<br/>
10.10.2.3 [Resume Statement](statements.md)<br/>
10.11 [Branch Statements](statements.md)<br/>
10.12 [Array-Handling Statements](statements.md)<br/>
10.12.1 [ReDim Statement](statements.md)<br/>
10.12.2 [Erase Statement](statements.md)<br/>
10.13 [Using statement](statements.md)<br/>
10.14 [Await Statements](statements.md)<br/>
10.15 [Yield Statements](statements.md)<br/>
11. [Expressions](expressions.md)<br/>
11.1 [Expression Classifications](expressions.md)<br/>
11.1.1 [Expression Reclassification](expressions.md)<br/>
11.2 [Constant Expressions](expressions.md)<br/>
11.3 [Late-Bound Expressions](expressions.md)<br/>
11.4 [Simple Expressions](expressions.md)<br/>
11.4.1 [Literal Expressions](expressions.md)<br/>
11.4.2 [Parenthesized Expressions](expressions.md)<br/>
11.4.3 [Instance Expressions](expressions.md)<br/>
11.4.4 [Simple Name Expressions](expressions.md)<br/>
11.4.5 [AddressOf Expressions](expressions.md)<br/>
11.5 [Type Expressions](expressions.md)<br/>
11.5.1 [GetType Expressions](expressions.md)<br/>
11.5.2 [TypeOf...Is Expressions](expressions.md)<br/>
11.5.3 [Is Expressions](expressions.md)<br/>
11.5.4 [GetXmlNamespace Expressions](expressions.md)<br/>
11.6 [Member Access Expressions](expressions.md)<br/>
11.6.1 [Identical Type and Member Names](expressions.md)<br/>
11.6.2 [Default Instances](expressions.md)<br/>
11.6.2.1 [Default Instances and Type Names](expressions.md)<br/>
11.6.2.2 [Group Classes](expressions.md)<br/>
11.6.3 [Extension Method Collection](expressions.md)<br/>
11.7 [Dictionary Member Access Expressions](expressions.md)<br/>
11.8 [Invocation Expressions](expressions.md)<br/>
11.8.1 [Overloaded Method Resolution](expressions.md)<br/>
11.8.1.1 [Specificity of members/types given an argument list](expressions.md)<br/>
11.8.1.2 [Genericity](expressions.md)<br/>
11.8.1.3 [Depth of genericity](expressions.md)<br/>
11.8.2 [Applicability To Argument List](expressions.md)<br/>
11.8.3 [Passing Arguments, and Picking Arguments for Optional Parameters](expressions.md)<br/>
11.8.4 [Conditional Methods](expressions.md)<br/>
11.8.5 [Type Argument Inference](expressions.md)<br/>
11.9 [Index Expressions](expressions.md)<br/>
11.10 [New Expressions](expressions.md)<br/>
11.10.1 [Object-Creation Expressions](expressions.md)<br/>
11.10.2 [Array Expressions](expressions.md)<br/>
11.10.2.1 [Array creation expressions](expressions.md)<br/>
11.10.2.2 [Array Literals](expressions.md)<br/>
11.10.3 [Delegate-Creation Expressions](expressions.md)<br/>
11.10.4 [Anonymous Object-Creation Expressions](expressions.md)<br/>
11.11 [Cast Expressions](expressions.md)<br/>
11.12 [Operator Expressions](expressions.md)<br/>
11.12.1 [Operator Precedence and Associativity](expressions.md)<br/>
11.12.2 [Object Operands](expressions.md)<br/>
11.12.3 [Operator Resolution](expressions.md)<br/>
11.13 [Arithmetic Operators](expressions.md)<br/>
11.13.1 [Unary Plus Operator](expressions.md)<br/>
11.13.2 [Unary Minus Operator](expressions.md)<br/>
11.13.3 [Addition Operator](expressions.md)<br/>
11.13.4 [Subtraction Operator](expressions.md)<br/>
11.13.5 [Multiplication Operator](expressions.md)<br/>
11.13.6 [Division Operators](expressions.md)<br/>
11.13.7 [Mod Operator](expressions.md)<br/>
11.13.8 [Exponentiation Operator](expressions.md)<br/>
11.14 [Relational Operators](expressions.md)<br/>
11.15 [Like Operator](expressions.md)<br/>
11.16 [Concatenation Operator](expressions.md)<br/>
11.17 [Logical Operators](expressions.md)<br/>
11.17.1 [Short-circuiting Logical Operators](expressions.md)<br/>
11.18 [Shift Operators](expressions.md)<br/>
11.19 [Boolean Expressions](expressions.md)<br/>
11.20 [Lambda Expressions](expressions.md)<br/>
11.20.1 [Closures](expressions.md)<br/>
11.21 [Query Expressions](expressions.md)<br/>
11.21.1 [Range Variables](expressions.md)<br/>
11.21.2 [Queryable Types](expressions.md)<br/>
11.21.3 [Default Query Indexer](expressions.md)<br/>
11.21.4 [From Query Operator](expressions.md)<br/>
11.21.5 [Join Query Operator](expressions.md)<br/>
11.21.6 [Let Query Operator](expressions.md)<br/>
11.21.7 [Select Query Operator](expressions.md)<br/>
11.21.8 [Distinct Query Operator](expressions.md)<br/>
11.21.9 [Where Query Operator](expressions.md)<br/>
11.21.10 [Partition Query Operators](expressions.md)<br/>
11.21.11 [Order By Query Operator](expressions.md)<br/>
11.21.12 [Group By Query Operator](expressions.md)<br/>
11.21.13 [Aggregate Query Operator](expressions.md)<br/>
11.21.14 [Group Join Query Operator](expressions.md)<br/>
11.22 [Conditional Expressions](expressions.md)<br/>
11.23 [XML Literal Expressions](expressions.md)<br/>
11.23.1 [Lexical rules](expressions.md)<br/>
11.23.2 [Embedded expressions](expressions.md)<br/>
11.23.3 [XML Documents](expressions.md)<br/>
11.23.4 [XML Elements](expressions.md)<br/>
11.23.5 [XML Namespaces](expressions.md)<br/>
11.23.6 [XML Processing Instructions](expressions.md)<br/>
11.23.7 [XML Comments](expressions.md)<br/>
11.23.8 [CDATA sections](expressions.md)<br/>
11.24 [XML Member Access Expressions](expressions.md)<br/>
11.25 [Await Operator](expressions.md)<br/>
12. [Documentation Comments](documentation-comments.md)<br/>
12.1 [Documentation Comment Format](documentation-comments.md)<br/>
12.2 [Recommended tags](documentation-comments.md)<br/>
12.2.1 [\<c\>](documentation-comments.md)<br/>
12.2.2 [\<code\>](documentation-comments.md)<br/>
12.2.3 [\<example\>](documentation-comments.md)<br/>
12.2.4 [\<exception\>](documentation-comments.md)<br/>
12.2.5 [\<include\>](documentation-comments.md)<br/>
12.2.6 [\<list\>](documentation-comments.md)<br/>
12.2.7 [\<para\>](documentation-comments.md)<br/>
12.2.8 [\<param\>](documentation-comments.md)<br/>
12.2.9 [\<paramref\>](documentation-comments.md)<br/>
12.2.10 [\<permission\>](documentation-comments.md)<br/>
12.2.11 [\<remarks\>](documentation-comments.md)<br/>
12.2.12 [\<returns\>](documentation-comments.md)<br/>
12.2.13 [\<see\>](documentation-comments.md)<br/>
12.2.14 [\<seealso\>](documentation-comments.md)<br/>
12.2.15 [\<summary\>](documentation-comments.md)<br/>
12.2.16 [\<typeparam\>](documentation-comments.md)<br/>
12.2.17 [\<value\>](documentation-comments.md)<br/>
12.3 [ID Strings](documentation-comments.md)<br/>
12.3.1 [ID string examples](documentation-comments.md)<br/>
12.4 [Documentation comments example](documentation-comments.md)<br/>
13. [Grammar Summary](grammar.md)<br/>
13.1 [Lexical Grammar](grammar.md)<br/>
13.1.1 [Characters and Lines](grammar.md)<br/>
13.1.2 [Identifiers](grammar.md)<br/>
13.1.3 [Keywords](grammar.md)<br/>
13.1.4 [Literals](grammar.md)<br/>
13.2 [Preprocessing Directives](grammar.md)<br/>
13.2.1 [Conditional Compilation](grammar.md)<br/>
13.2.2 [External Source Directives](grammar.md)<br/>
13.2.3 [Region Directives](grammar.md)<br/>
13.2.4 [External Checksum Directives](grammar.md)<br/>
13.3 [Syntactic Grammar](grammar.md)<br/>
13.3.1 [Attributes](grammar.md)<br/>
13.3.2 [Source Files and Namespaces](grammar.md)<br/>
13.3.3 [Types](grammar.md)<br/>
13.3.4 [Type Members](grammar.md)<br/>
13.3.5 [Statements](grammar.md)<br/>
13.3.6 [Expressions](grammar.md)<br/>
14. [Change List](change-list.md)<br/>
14.1 [Version 7.1 to Version 8.0](change-list.md)<br/>
14.1.1 [Major changes](change-list.md)<br/>
14.1.2 [Minor changes](change-list.md)<br/>
14.1.3 [Clarifications/Errata](change-list.md)<br/>
14.1.4 [Miscellaneous](change-list.md)<br/>
14.2 [Version 8.0 to Version 8.0 (2nd Edition)](change-list.md)<br/>
14.2.1 [Minor changes](change-list.md)<br/>
14.2.2 [Clarifications/Errata](change-list.md)<br/>
14.2.3 [Miscellaneous](change-list.md)<br/>
14.3 [Version 8.0 (2nd Edition) to Version 9.0](change-list.md)<br/>
14.3.1 [Major Changes](change-list.md)<br/>
14.3.2 [Minor Changes](change-list.md)<br/>
14.3.3 [Clarifications/Errata](change-list.md)<br/>
14.3.4 [Miscellaneous](change-list.md)<br/>
14.4 [Version 9.0 to Version 10.0](change-list.md)<br/>
14.4.1 [Major Changes](change-list.md)<br/>
14.4.2 [Minor Changes](change-list.md)<br/>
14.4.3 [Clarifications/Errata](change-list.md)<br/>
14.5 [Version 10.0 to Version 11.0](change-list.md)<br/>
14.5.1 [Major Changes](change-list.md)<br/>
14.5.2 [Minor Changes](change-list.md)<br/>
14.5.3 [Clarifications/Errata](change-list.md)<br/>

