# Grammar Summary

This section summarizes the Visual Basic language grammar. For information on how to read the grammar, see Grammar Notation.

## Lexical Grammar

<pre>Start  ::=  [  LogicalLine+  ]</pre>

<pre>LogicalLine  ::=  [  LogicalLineElement+  ]  [  Comment  ]  LineTerminator</pre>

<pre>LogicalLineElement  ::=  WhiteSpace  |  LineContinuation  |  Token</pre>

<pre>Token  ::=  Identifier  |  Keyword  |  Literal  |  Separator  |  Operator</pre>

### Characters and Lines

<pre>Character  ::=  < any Unicode character except a *LineTerminator* ></pre>

<pre>LineTerminator  ::=
    < Unicode carriage return character (0x000D) >  |
    < Unicode linefeed character (0x000A) >  |
    < Unicode carriage return character >  < Unicode linefeed character >  |
    < Unicode line separator character (0x2028) >  |
    < Unicode paragraph separator character (0x2029) ></pre>

<pre>LineContinuation  ::=  WhiteSpace  <b>_</b>  [  WhiteSpace+  ]  LineTerminator</pre>

<pre>Comma  ::=  <b>,</b>  [  LineTerminator  ]</pre>

<pre>OpenParenthesis  ::=  <b>(</b>  [  LineTerminator  ]</pre>

<pre>CloseParenthesis  ::=  [  LineTerminator  ]  <b>)</b></pre>

<pre>OpenCurlyBrace  ::=  <b>{</b>  [  LineTerminator  ]</pre>

<pre>CloseCurlyBrace  ::=  [  LineTerminator  ]  <b>}</b></pre>

<pre>Equals  ::=  <b>=</b>  [  LineTerminator  ]</pre>

<pre>ColonEquals  ::=  <b>:</b><b>=</b>  [  LineTerminator  ]</pre>

<pre>WhiteSpace  ::=
    < Unicode blank characters (class Zs) >  |
    < Unicode tab character (0x0009) ></pre>

<pre>Comment  ::=  CommentMarker  [  Character+  ]</pre>

<pre>CommentMarker  ::=  SingleQuoteCharacter  |  <b>REM</b></pre>

<pre>SingleQuoteCharacter  ::=
<b>'</b>  |
    < Unicode left single-quote character (0x2018) >  |
    < Unicode right single-quote character (0x2019) ></pre>

### Identifiers

<pre>Identifier  ::=
    NonEscapedIdentifier  [  TypeCharacter  ]  |
    Keyword  TypeCharacter  |
    EscapedIdentifier</pre>

<pre>NonEscapedIdentifier  ::=  < *IdentifierName* but not *Keyword* ></pre>

<pre>EscapedIdentifier  ::=  <b>[</b>  IdentifierName  <b>]</b></pre>

<pre>IdentifierName  ::=  IdentifierStart  [  IdentifierCharacter+  ]</pre>

<pre>IdentifierStart  ::=
    AlphaCharacter  |
    UnderscoreCharacter  IdentifierCharacter</pre>

<pre>IdentifierCharacter  ::=
    UnderscoreCharacter  |
    AlphaCharacter  |
    NumericCharacter  |
    CombiningCharacter  |
    FormattingCharacter</pre>

<pre>AlphaCharacter  ::=
    < Unicode alphabetic character (classes Lu, Ll, Lt, Lm, Lo, Nl) ></pre>

<pre>NumericCharacter  ::=  < Unicode decimal digit character (class Nd) ></pre>

<pre>CombiningCharacter  ::=  < Unicode combining character (classes Mn, Mc) ></pre>

<pre>FormattingCharacter  ::=  < Unicode formatting character (class Cf) ></pre>

<pre>UnderscoreCharacter  ::=  < Unicode connection character (class Pc) ></pre>

<pre>IdentifierOrKeyword  ::=  Identifier  |  Keyword</pre>

<pre>TypeCharacter  ::=
    IntegerTypeCharacter  |
    LongTypeCharacter  |
    DecimalTypeCharacter  |
    SingleTypeCharacter  |
    DoubleTypeCharacter  |
    StringTypeCharacter</pre>

<pre>IntegerTypeCharacter  ::=  <b>%</b></pre>

<pre>LongTypeCharacter  ::=  <b>&</b></pre>

<pre>DecimalTypeCharacter  ::=  <b>@</b></pre>

<pre>SingleTypeCharacter  ::=  <b>!</b></pre>

<pre>DoubleTypeCharacter  ::=  <b>#</b></pre>

<pre>StringTypeCharacter  ::=  <b>$</b></pre>

### Keywords

<pre>Keyword  ::=  < member of keyword table in 2.3 ></pre>

### Literals

<pre>Literal  ::=
    BooleanLiteral  |
    IntegerLiteral  |
    FloatingPointLiteral  |
    StringLiteral  |
    CharacterLiteral  |
    DateLiteral  |
    Nothing</pre>

<pre>BooleanLiteral  ::=  <b>True</b>  |  <b>False</b></pre>

<pre>IntegerLiteral  ::=  IntegralLiteralValue  [  IntegralTypeCharacter  ]</pre>

<pre>IntegralLiteralValue  ::=  IntLiteral  |  HexLiteral  |  OctalLiteral</pre>

<pre>IntegralTypeCharacter  ::=
    ShortCharacter  |
    UnsignedShortCharacter  |
    IntegerCharacter  |
    UnsignedIntegerCharacter  |
    LongCharacter  |
    UnsignedLongCharacter  |
    IntegerTypeCharacter  |
    LongTypeCharacter</pre>

<pre>ShortCharacter  ::=  <b>S</b></pre>

<pre>UnsignedShortCharacter  ::=  <b>US</b></pre>

<pre>IntegerCharacter  ::=  <b>I</b></pre>

<pre>UnsignedIntegerCharacter  ::=  <b>UI</b></pre>

<pre>LongCharacter  ::=  <b>L</b></pre>

<pre>UnsignedLongCharacter  ::=  <b>UL</b></pre>

<pre>IntLiteral  ::=  Digit+</pre>

<pre>HexLiteral  ::=  <b>&</b><b>H</b>  HexDigit+</pre>

<pre>OctalLiteral  ::=  <b>&</b><b>O</b>  OctalDigit+</pre>

<pre>Digit  ::=  <b>0</b>  |  <b>1</b>  |  <b>2</b>  |  <b>3</b>  |  <b>4</b>  |  <b>5</b>  |  <b>6</b>  |  <b>7</b>  |  <b>8</b>  |  <b>9</b></pre>

<pre>HexDigit  ::=  <b>0</b>  |  <b>1</b>  |  <b>2</b>  |  <b>3</b>  |  <b>4</b>  |  <b>5</b>  |  <b>6</b>  |  <b>7</b>  |  <b>8</b>  |  <b>9</b>  |  <b>A</b>  |  <b>B</b>  |  <b>C</b>  |  <b>D</b>  |  <b>E</b>  |  <b>F</b></pre>

<pre>OctalDigit  ::=  <b>0</b>  |  <b>1</b>  |  <b>2</b>  |  <b>3</b>  |  <b>4</b>  |  <b>5</b>  |  <b>6</b>  |  <b>7</b></pre>

<pre>FloatingPointLiteral  ::=
    FloatingPointLiteralValue  [  FloatingPointTypeCharacter  ]  |
    IntLiteral  FloatingPointTypeCharacter</pre>

<pre>FloatingPointTypeCharacter  ::=
    SingleCharacter  |
    DoubleCharacter  |
    DecimalCharacter  |
    SingleTypeCharacter  |
    DoubleTypeCharacter  |
    DecimalTypeCharacter</pre>

<pre>SingleCharacter  ::=  <b>F</b></pre>

<pre>DoubleCharacter  ::=  <b>R</b></pre>

<pre>DecimalCharacter  ::=  <b>D</b></pre>

<pre>FloatingPointLiteralValue  ::=
    IntLiteral  <b>.</b>  IntLiteral  [  Exponent  ]  |
<b>.</b>  IntLiteral  [  Exponent  ]  |
    IntLiteral  Exponent</pre>

<pre>Exponent  ::=  <b>E</b>  [  Sign  ]  IntLiteral</pre>

<pre>Sign  ::=  <b>+</b>  |  <b>-</b></pre>

<pre>StringLiteral  ::=
    DoubleQuoteCharacter  [  StringCharacter+  ]  DoubleQuoteCharacter</pre>

<pre>DoubleQuoteCharacter  ::=
<b>"</b>  |
    < Unicode left double-quote character (0x201C) >  |
    < Unicode right double-quote character (0x201D) ></pre>

<pre>StringCharacter  ::=
    < Character except for DoubleQuoteCharacter >  |
    DoubleQuoteCharacter  DoubleQuoteCharacter</pre>

<pre>CharacterLiteral  ::=  DoubleQuoteCharacter  StringCharacter  DoubleQuoteCharacter  <b>C</b></pre>

<pre>DateLiteral  ::=  <b>#</b>  [  Whitespace+  ]  DateOrTime  [  Whitespace+  ]  <b>#</b></pre>

<pre>DateOrTime  ::=
    DateValue  Whitespace+  TimeValue  |
    DateValue  |
    TimeValue</pre>

<pre>DateValue  ::=
    MonthValue  <b>/</b>  DayValue  <b>/</b>  YearValue  |
    MonthValue  <b>-</b>  DayValue  <b>-</b>  YearValue</pre>

<pre>TimeValue  ::=
    HourValue  <b>:</b>  MinuteValue  [  <b>:</b>  SecondValue  ]  [  WhiteSpace+  ]  [  AMPM  ]  |
    HourValue  [  WhiteSpace+  ]  AMPM</pre>

<pre>MonthValue  ::=  IntLiteral</pre>

<pre>DayValue  ::=  IntLiteral</pre>

<pre>YearValue  ::=  IntLiteral</pre>

<pre>HourValue  ::=  IntLiteral</pre>

<pre>MinuteValue  ::=  IntLiteral</pre>

<pre>SecondValue  ::=  IntLiteral</pre>

<pre>AMPM  ::=  <b>AM</b>  |  <b>PM</b></pre>

<pre>ElseIf ::=  <b>ElseIf</b>  |  <b>Else If</b></pre>

<pre>Nothing  ::=  <b>Nothing</b></pre>

<pre>Separator  ::=  <b>(</b>  |  <b>)</b>  |  <b>{</b>  |  <b>}</b>  |  <b>!</b>  |  <b>#</b>  |  <b>,</b>  |  <b>.</b>  |  <b>:</b>  |  <b>?</b></pre>

<pre>Operator  ::=
<b>&</b>  |  <b>*</b>  |  <b>+</b>  |  <b>-</b>  |  <b>/</b>  |  <b>\</b>  |  <b>^</b>  |  <b><</b>  |  <b>=</b>  |  <b>></b></pre>

## Preprocessing Directives

### Conditional Compilation

<pre>Start  ::=  [  CCStatement+  ]</pre>

<pre>CCStatement  ::=
    CCConstantDeclaration  |
    CCIfGroup  |
    LogicalLine</pre>

<pre>CCExpression  ::=
    LiteralExpression  |
    CCParenthesizedExpression  |
    CCSimpleNameExpression  |
    CCCastExpression  |
    CCOperatorExpression  |
    CCConditionalExpression</pre>

<pre>CCParenthesizedExpression  ::=  <b>(</b>  CCExpression  <b>)</b></pre>

<pre>CCSimpleNameExpression  ::=  Identifier</pre>

<pre>CCCastExpression  ::=  
<b>DirectCast</b><b>(</b>  CCExpression  <b>,</b>  TypeName  <b>)</b>  |
<b>TryCast</b><b>(</b>  CCExpression  <b>,</b>  TypeName  <b>)</b>  |
<b>CType</b><b>(</b>  CCExpression  <b>,</b>  TypeName  <b>)</b>  |
    CastTarget  <b>(</b>  CCExpression  <b>)</b></pre>

<pre>CCOperatorExpression  ::=
    CCUnaryOperator  CCExpression  |
    CCExpression  CCBinaryOperator  CCExpression</pre>

<pre>CCUnaryOperator  ::=  <b>+</b>  |  <b>-</b>  |  <b>Not</b></pre>

<pre>CCBinaryOperator  ::=  <b>+</b>  |  <b>-</b>  |  <b>*</b>  |  <b>/</b>  |  <b>\</b>  |  <b>Mod</b>  |  <b>^</b>  |  <b>=</b>  |  <b><</b><b>></b>  |  <b><</b>  |  <b>></b>  |
<b><</b><b>=</b>  |  <b>></b><b>=</b>  |  <b>&</b>  |  <b>And</b>  |  <b>Or</b>  |  <b>Xor</b>  |  <b>AndAlso</b>  |  <b>OrElse</b>  |  <b><</b><b><</b>  |  <b>></b><b>></b></pre>

<pre>CCConditionalExpression  ::=  
<b>If</b><b>(</b>  CCExpression  <b>,</b>  CCExpression  <b>,</b>  CCExpression  <b>)</b>  |
<b>If</b><b>(</b>  CCExpression  <b>,</b>  CCExpression  <b>)</b></pre>

<pre>CCConstantDeclaration  ::=  <b>#</b><b>Const</b>  Identifier  <b>=</b>  CCExpression  LineTerminator</pre>

<pre>CCIfGroup  ::=
<b>#</b><b>If</b>  CCExpression  [  <b>Then</b>  ]  LineTerminator
    [  CCStatement+  ]
    [  CCElseIfGroup+  ]
    [  CCElseGroup  ]
<b>#</b><b>End</b><b>If</b>  LineTerminator</pre>

<pre>CCElseIfGroup  ::=
<b>#</b>  ElseIf  CCExpression  [  <b>Then</b>  ]  LineTerminator
    [  CCStatement+  ]</pre>

<pre>CCElseGroup  ::=
<b>#</b><b>Else</b>  LineTerminator
    [  CCStatement+  ]</pre>

### External Source Directives

<pre>Start  ::=  [  ExternalSourceStatement+  ]</pre>

<pre>ExternalSourceStatement  ::=  ExternalSourceGroup  |  LogicalLine</pre>

<pre>ExternalSourceGroup  ::=
<b>#</b><b>ExternalSource</b><b>(</b>  StringLiteral  <b>,</b>  IntLiteral  <b>)</b>  LineTerminator
    [  LogicalLine+  ]
<b>#</b><b>End</b><b>ExternalSource</b>  LineTerminator</pre>

### Region Directives

<pre>Start  ::=  [  RegionStatement+  ]</pre>

<pre>RegionStatement  ::=  RegionGroup  |  LogicalLine</pre>

<pre>RegionGroup  ::=
<b>#</b><b>Region</b>  StringLiteral  LineTerminator
    [  RegionStatement+  ]
<b>#</b><b>End</b><b>Region</b>  LineTerminator</pre>

### External Checksum Directives

<pre>Start  ::=  [  ExternalChecksumStatement+  ]</pre>

<pre>ExternalChecksumStatement  ::=
<b>#</b><b>ExternalChecksum</b><b>(</b>  StringLiteral  <b>,</b>  StringLiteral  <b>,</b>  StringLiteral  <b>)</b>  LineTerminator</pre>

## Syntactic Grammar

<pre>AccessModifier  ::=  <b>Public</b>  |  <b>Protected</b>  |  <b>Friend</b>  |  <b>Private</b>  |  <b>Protected</b><b>Friend</b></pre>

<pre>TypeParameterList  ::=
    OpenParenthesis  <b>Of</b>  TypeParameters  CloseParenthesis</pre>

<pre>TypeParameters  ::=
    TypeParameter  |
    TypeParameters  Comma  TypeParameter</pre>

<pre>TypeParameter  ::=
    [  VarianceModifier  ]  Identifier  [  TypeParameterConstraints  ]</pre>

<pre>VarianceModifier  ::=
<b>In</b>  |  <b>Out</b></pre>

<pre>TypeParameterConstraints  ::=
<b>As</b>  Constraint  |
<b>As</b>  OpenCurlyBrace  ConstraintList  CloseCurlyBrace</pre>

<pre>ConstraintList  ::=
    ConstraintList  *Comma*  Constraint  |
    Constraint</pre>

<pre>Constraint  ::=  TypeName  |  <b>New</b>  |  <b>Structure</b>  |  <b>Class</b></pre>

### Attributes

<pre>Attributes  ::=
    AttributeBlock  |
    Attributes  AttributeBlock</pre>

<pre>AttributeBlock  ::=  [  LineTerminator  ]  <b><</b>  AttributeList  [  LineTerminator  ]  <b>></b>  [  LineTerminator  ]</pre>

<pre>AttributeList  ::=
    Attribute  |
    AttributeList  *Comma*  Attribute</pre>

<pre>Attribute  ::=
    [  AttributeModifier  <b>:</b>  ]  SimpleTypeName  [  OpenParenthesis  [  AttributeArguments  ]  CloseParenthesis  ]</pre>

<pre>AttributeModifier  ::=  <b>Assembly</b>  |  <b>Module</b></pre>

<pre>AttributeArguments  ::=
    AttributePositionalArgumentList  |
    AttributePositionalArgumentList  Comma  VariablePropertyInitializerList  |
    VariablePropertyInitializerList</pre>

<pre>AttributePositionalArgumentList  ::=
    AttributeArgumentExpression  |
    AttributePositionalArgumentList  Comma  AttributeArgumentExpression</pre>

<pre>VariablePropertyInitializerList  ::=
    VariablePropertyInitializer  |
    VariablePropertyInitializerList  Comma  VariablePropertyInitializer</pre>

<pre>VariablePropertyInitializer  ::=
    IdentifierOrKeyword  ColonEquals  AttributeArgumentExpression</pre>

<pre>AttributeArgumentExpression  ::=
    ConstantExpression  |
    GetTypeExpression  |
    ArrayExpression</pre>

### Source Files and Namespaces

<pre>Start  ::=
    [  OptionStatement+  ]
    [  ImportsStatement+  ]
    [  AttributesStatement+  ]
    [  NamespaceMemberDeclaration+  ]</pre>

<pre>StatementTerminator  ::=  LineTerminator  |  <b>:</b></pre>

<pre>AttributesStatement  ::=  Attributes  StatementTerminator</pre>

<pre>OptionStatement  ::=
    OptionExplicitStatement  |
    OptionStrictStatement  |
    OptionCompareStatement  |
    OptionInferStatement</pre>

<pre>OptionExplicitStatement  ::=  <b>Option</b><b>Explicit</b>  [  OnOff  ]  StatementTerminator</pre>

<pre>OnOff  ::=  <b>On</b>  |  <b>Off</b></pre>

<pre>OptionStrictStatement  ::=  <b>Option</b><b>Strict</b>  [  OnOff  ]  StatementTerminator</pre>

<pre>OptionCompareStatement  ::=  <b>Option</b><b>Compare</b>  CompareOption  StatementTerminator</pre>

<pre>CompareOption  ::=  <b>Binary</b>  |  <b>Text</b></pre>

<pre>OptionInferStatement  ::=  <b>Option</b><b>Infer</b>  [  OnOff  ]  StatementTerminator</pre>

<pre>ImportsStatement  ::=  <b>Imports</b>  ImportsClauses  StatementTerminator</pre>

<pre>ImportsClauses  ::=
    ImportsClause  |
    ImportsClauses  Comma  ImportsClause</pre>

<pre>ImportsClause  ::=
    AliasImportsClause  |
    MembersImportsClause  |
    XMLNamespaceImportsClause</pre>

<pre>AliasImportsClause  ::=  
    Identifier  Equals  TypeName</pre>

<pre>MembersImportsClause  ::=
    TypeName</pre>

<pre>XMLNamespaceImportsClause  ::=
<b><</b>  XMLNamespaceAttributeName  [  XMLWhitespace  ]  *Equals*   [  XMLWhitespace  ]  XMLNamespaceValue  <b>></b></pre>

<pre>XMLNamespaceValue  ::=
    DoubleQuoteCharacter  [  XMLAttributeDoubleQuoteValueCharacter+  ]  DoubleQuoteCharacter  |
    SingleQuoteCharacter  [  XMLAttributeSingleQuoteValueCharacter+  ]  SingleQuoteCharacter</pre>

<pre>NamespaceDeclaration  ::=
<b>Namespace</b>  NamespaceName  StatementTerminator
    [  NamespaceMemberDeclaration+  ]
<b>End</b><b>Namespace</b>  StatementTerminator</pre>

<pre>NamespaceName  ::= 
    RelativeNamespaceName  |
<b>Global</b>  |
<b>Global</b>  .   RelativeNamespaceName</pre>

<pre>*RelativeNamespaceName  ::=*
    Identifier  |
*Relative*NamespaceName  Period  IdentifierOrKeyword</pre>

<pre>NamespaceMemberDeclaration  ::=
    NamespaceDeclaration  |
    TypeDeclaration</pre>

<pre>TypeDeclaration  ::=
    ModuleDeclaration  |
    NonModuleDeclaration</pre>

<pre>NonModuleDeclaration  ::=
    EnumDeclaration  |
    StructureDeclaration  |
    InterfaceDeclaration  |
    ClassDeclaration  |
    DelegateDeclaration</pre>

### Types

<pre>TypeName  ::=
    ArrayTypeName  |
    NonArrayTypeName</pre>

<pre>NonArrayTypeName  ::=
    SimpleTypeName  |    NullableTypeName</pre>

<pre>SimpleTypeName  ::=
    QualifiedTypeName  |
    BuiltInTypeName</pre>

<pre>QualifiedTypeName  ::=
    Identifier  [  TypeArguments  ]  |
<b>Global</b>  Period  IdentifierOrKeyword    [  TypeArguments  ]  |
    QualifiedTypeName  Period  IdentifierOrKeyword  [  TypeArguments  ]</pre>

<pre>TypeArguments  ::=
    OpenParenthesis  <b>Of</b>  TypeArgumentList  CloseParenthesis</pre>

<pre>TypeArgumentList  ::=
    TypeName  |
    TypeArgumentList  Comma  TypeName</pre>

<pre>BuiltInTypeName  ::=  <b>Object</b>  |  PrimitiveTypeName</pre>

<pre>TypeModifier  ::=  AccessModifier  |  <b>Shadows</b></pre>

<pre>IdentifierModifiers  ::=  [ *NullableNameModifier* ]  [ ArrayNameModifier  ]</pre>

<pre>NullableTypeName  ::=  NonArrayTypeName  <b>?</b></pre>

<pre>NullableNameModifier  ::=  <b>?</b></pre>

<pre>TypeImplementsClause  ::=  <b>Implements</b>*Type*Implements  StatementTerminator</pre>

<pre>TypeImplements  ::=
    NonArrayTypeName  |
*Type*Implements  Comma  NonArrayTypeName</pre>

<pre>PrimitiveTypeName  ::=  NumericTypeName  |  <b>Boolean</b>  |  <b>Date</b>  |  <b>Char</b>  |  <b>String</b></pre>

<pre>NumericTypeName  ::=  IntegralTypeName  |  FloatingPointTypeName  |  <b>Decimal</b></pre>

<pre>IntegralTypeName  ::=  <b>Byte</b>  |  <b>SByte</b>  |  <b>UShort</b>  |  <b>Short</b>  |  <b>UInteger</b>  |  <b>Integer</b>  |  <b>ULong</b>  |  <b>Long</b></pre>

<pre>FloatingPointTypeName  ::=  <b>Single</b>  |  <b>Double</b></pre>

<pre>EnumDeclaration  ::=
    [  Attributes  ]  [  TypeModifier+  ]  <b>Enum</b>  Identifier  [  <b>As</b>  NonArrayTypeName  ]  StatementTerminator
    EnumMemberDeclaration+
<b>End</b><b>Enum</b>  StatementTerminator</pre>

<pre>EnumMemberDeclaration  ::=  [  Attributes  ]  Identifier  [  Equals  ConstantExpression  ]  StatementTerminator</pre>

<pre>ClassDeclaration  ::=
    [  Attributes  ]  [  ClassModifier+  ]  <b>Class</b>  Identifier  [  TypeParameterList  ]  StatementTerminator
    [  ClassBase  ]
    [  TypeImplementsClause+  ]
    [  ClassMemberDeclaration+  ]
<b>End</b><b>Class</b>  StatementTerminator</pre>

<pre>ClassModifier  ::=  TypeModifier  |  <b>MustInherit</b>  |  <b>NotInheritable</b>  |  <b>Partial</b></pre>

<pre>ClassBase  ::=  <b>Inherits</b>  NonArrayTypeName  StatementTerminator</pre>

<pre>ClassMemberDeclaration  ::=
    NonModuleDeclaration  |
    EventMemberDeclaration  |
    VariableMemberDeclaration  |
    ConstantMemberDeclaration  |
    MethodMemberDeclaration  |
    PropertyMemberDeclaration  |
    ConstructorMemberDeclaration  |
    OperatorDeclaration</pre>

<pre>StructureDeclaration  ::=
    [  Attributes  ]  [  StructureModifier+  ]  <b>Structure</b>  Identifier  [  TypeParameterList  ]
        StatementTerminator
    [  TypeImplementsClause+  ]
    [  StructMemberDeclaration+  ]
<b>End</b><b>Structure</b>  StatementTerminator</pre>

<pre>StructureModifier  ::=  TypeModifier  |  <b>Partial</b></pre>

<pre>StructMemberDeclaration  ::=
    NonModuleDeclaration  |
    VariableMemberDeclaration  |
    ConstantMemberDeclaration  |
    EventMemberDeclaration  |
    MethodMemberDeclaration  |
    PropertyMemberDeclaration  |
    ConstructorMemberDeclaration  |
    OperatorDeclaration</pre>

<pre>ModuleDeclaration  ::=
    [  Attributes  ]  [  TypeModifier+  ]  <b>Module</b>  Identifier  StatementTerminator
    [  ModuleMemberDeclaration+  ]
<b>End</b><b>Module</b>  StatementTerminator</pre>

<pre>ModuleMemberDeclaration  ::=
    NonModuleDeclaration  |
    VariableMemberDeclaration  |
    ConstantMemberDeclaration  |
    EventMemberDeclaration  |
    MethodMemberDeclaration  |
    PropertyMemberDeclaration  |
    ConstructorMemberDeclaration</pre>

<pre>InterfaceDeclaration  ::=
    [  Attributes  ]  [  TypeModifier+  ]  <b>Interface</b>  Identifier  [  TypeParameterList  ]  StatementTerminator
    [  InterfaceBase+  ]
    [  InterfaceMemberDeclaration+  ]
<b>End</b><b>Interface</b>  StatementTerminator</pre>

<pre>InterfaceBase  ::=  <b>Inherits</b>  InterfaceBases  StatementTerminator</pre>

<pre>InterfaceBases  ::=
    NonArrayTypeName  |
    InterfaceBases  Comma  NonArrayTypeName</pre>

<pre>InterfaceMemberDeclaration  ::=
    NonModuleDeclaration  |
    InterfaceEventMemberDeclaration  |
    InterfaceMethodMemberDeclaration  |
    InterfacePropertyMemberDeclaration</pre>

<pre>ArrayTypeName  ::=  NonArrayTypeName  ArrayTypeModifiers</pre>

<pre>ArrayTypeModifiers  ::=  ArrayTypeModifier+</pre>

<pre>ArrayTypeModifier  ::=  OpenParenthesis  [  RankList  ]  CloseParenthesis</pre>

<pre>RankList  ::=
    Comma  |
    RankList  Comma</pre>

<pre>ArrayNameModifier  ::=
    ArrayTypeModifiers  |
    ArraySizeInitializationModifier</pre>

<pre>DelegateDeclaration  ::=
    [  Attributes  ]  [  TypeModifier+  ]  <b>Delegate</b>  MethodSignature  StatementTerminator</pre>

<pre>MethodSignature  ::=  SubSignature  |  FunctionSignature</pre>

### Type Members

<pre>ImplementsClause  ::=  [  <b>Implements</b>  ImplementsList  ]</pre>

<pre>ImplementsList  ::=
    InterfaceMemberSpecifier  |
    ImplementsList  Comma  InterfaceMemberSpecifier</pre>

<pre>InterfaceMemberSpecifier  ::=  NonArrayTypeName  Period  IdentifierOrKeyword</pre>

<pre>MethodMemberDeclaration  ::=  MethodDeclaration  |  ExternalMethodDeclaration</pre>

<pre>InterfaceMethodMemberDeclaration  ::=  InterfaceMethodDeclaration</pre>

<pre>MethodDeclaration  ::=
    SubDeclaration  |
    MustOverrideSubDeclaration  |
    FunctionDeclaration  |
    MustOverrideFunctionDeclaration</pre>

<pre>InterfaceMethodDeclaration  ::=
    InterfaceSubDeclaration  |
    InterfaceFunctionDeclaration</pre>

<pre>SubSignature  ::=  <b>Sub</b>  Identifier  [  TypeParameterList  ]
    [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]</pre>

<pre>FunctionSignature  ::=  <b>Function</b>  Identifier  [  TypeParameterList  ]
        [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]  [  <b>As</b>  [  Attributes  ]  TypeName  ]</pre>

<pre>SubDeclaration  ::=
    [  Attributes  ]  [  ProcedureModifier+  ]  SubSignature  [  HandlesOrImplements  ]  LineTerminator
    Block
<b>End</b><b>Sub</b>  StatementTerminator</pre>

<pre>MustOverrideSubDeclaration  ::=
    [  Attributes  ]  MustOverrideProcedureModifier+  SubSignature  [  HandlesOrImplements  ]
        StatementTerminator</pre>

<pre>InterfaceSubDeclaration  ::=
    [  Attributes  ]  [  InterfaceProcedureModifier+  ]  SubSignature  StatementTerminator</pre>

<pre>FunctionDeclaration  ::=
    [  Attributes  ]  [  ProcedureModifier+  ]  FunctionSignature  [  HandlesOrImplements  ]
        LineTerminator
    Block
<b>End</b><b>Function</b>  StatementTerminator</pre>

<pre>MustOverrideFunctionDeclaration  ::=
    [  Attributes  ]  MustOverrideProcedureModifier+  FunctionSignature
        [  HandlesOrImplements  ]  StatementTerminator</pre>

<pre>InterfaceFunctionDeclaration  ::=
    [  Attributes  ]  [  InterfaceProcedureModifier+  ]  FunctionSignature  StatementTerminator</pre>

<pre>ProcedureModifier  ::=
    AccessModifier  |
<b>Shadows</b>  |
<b>Shared</b>  |
<b>Overridable</b>  |
<b>NotOverridable</b>  |
<b>Overrides</b>  |
<b>Overloads</b>  |
<b>Partial</b>  |
<b>Iterator</b>  |
<b>Async</b></pre>

<pre>MustOverrideProcedureModifier  ::=  ProcedureModifier  |  <b>MustOverride</b></pre>

<pre>InterfaceProcedureModifier  ::=  <b>Shadows</b>  |  <b>Overloads</b></pre>

<pre>HandlesOrImplements  ::=  HandlesClause  |  ImplementsClause</pre>

<pre>ExternalMethodDeclaration  ::=
    ExternalSubDeclaration  |
    ExternalFunctionDeclaration</pre>

<pre>ExternalSubDeclaration  ::=
    [  Attributes  ]  [  ExternalMethodModifier+  ]  <b>Declare</b>  [  CharsetModifier  ]  <b>Sub</b>  Identifier
        LibraryClause  [  AliasClause  ]  [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]  StatementTerminator</pre>

<pre>ExternalFunctionDeclaration  ::=
    [  Attributes  ]  [  ExternalMethodModifier+  ]  <b>Declare</b>  [  CharsetModifier  ]  <b>Function</b>  Identifier
        LibraryClause  [  AliasClause  ]  [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]  [  <b>As</b>  [  Attributes  ]  TypeName  ]
        StatementTerminator</pre>

<pre>ExternalMethodModifier  ::=  AccessModifier  |  <b>Shadows</b>  |  <b>Overloads</b></pre>

<pre>CharsetModifier  ::=  <b>Ansi</b>  |  <b>Unicode</b>  |  <b>Auto</b></pre>

<pre>LibraryClause  ::=  <b>Lib</b>  StringLiteral</pre>

<pre>AliasClause  ::=  <b>Alias</b>  StringLiteral</pre>

<pre>ParameterList  ::=
    Parameter  |
    ParameterList  Comma  Parameter</pre>

<pre>Parameter  ::=
    [  Attributes  ]  [  ParameterModifier+  ]  ParameterIdentifier  [  <b>As</b>  TypeName  ]
        [  *Equals*  ConstantExpression  ]</pre>

<pre>ParameterModifier  ::=  <b>ByVal</b>  |  <b>ByRef</b>  |  <b>Optional</b>  |  <b>ParamArray</b></pre>

<pre>ParameterIdentifier  ::=  Identifier  IdentifierModifiers</pre>

<pre>HandlesClause  ::=  [  <b>Handles</b>  EventHandlesList  ]</pre>

<pre>EventHandlesList  ::=
    EventMemberSpecifier  |
    EventHandlesList  Comma  EventMemberSpecifier</pre>

<pre>EventMemberSpecifier  ::=
    Identifier  Period  IdentifierOrKeyword  |
<b>MyBase</b>  Period  IdentifierOrKeyword  |
<b>Me</b>  Period  IdentifierOrKeyword</pre>

<pre>ConstructorMemberDeclaration  ::=
    [  Attributes  ]  [  ConstructorModifier+  ]  <b>Sub</b><b>New</b>
        [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]  LineTerminator
    [  Block  ]
<b>End</b><b>Sub</b>  StatementTerminator</pre>

<pre>ConstructorModifier  ::=  AccessModifier  |  <b>Shared</b></pre>

<pre>EventMemberDeclaration  ::=
    RegularEventMemberDeclaration  |
    CustomEventMemberDeclaration</pre>

<pre>RegularEventMemberDeclaration  ::=
    [  Attributes  ]  [  EventModifiers+  ]  <b>Event</b>  Identifier  ParametersOrType  [  ImplementsClause  ]
        StatementTerminator</pre>

<pre>InterfaceEventMemberDeclaration  ::=
    [  Attributes  ]  [  InterfaceEventModifiers+  ]  <b>Event</b>  Identifier  ParametersOrType  StatementTerminator</pre>

<pre>ParametersOrType  ::=
    [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]  |
<b>As</b>  NonArrayTypeName</pre>

<pre>EventModifiers  ::=  AccessModifier  |  <b>Shadows</b>  |  <b>Shared</b></pre>

<pre>InterfaceEventModifiers  ::=  <b>Shadows</b></pre>

<pre>CustomEventMemberDeclaration  ::=
    [  Attributes  ]  [  EventModifiers+  ]  <b>Custom</b><b>Event</b>  Identifier  <b>As</b>  TypeName  [  ImplementsClause  ]
        StatementTerminator
        EventAccessorDeclaration+
<b>End</b><b>Event</b>  StatementTerminator</pre>

<pre>EventAccessorDeclaration  ::=
    AddHandlerDeclaration  |
    RemoveHandlerDeclaration  |
    RaiseEventDeclaration</pre>

<pre>AddHandlerDeclaration  ::=
    [  Attributes  ]  <b>AddHandler</b>  OpenParenthesis  ParameterList  CloseParenthesis  LineTerminator
    [  Block  ]
<b>End</b><b>AddHandler</b>  StatementTerminator</pre>

<pre>RemoveHandlerDeclaration  ::=
    [  Attributes  ]  <b>RemoveHandler</b>  OpenParenthesis  ParameterList  CloseParenthesis  LineTerminator
    [  Block  ]
<b>End</b><b>RemoveHandler</b>  StatementTerminator</pre>

<pre>RaiseEventDeclaration  ::=
    [  Attributes  ]  <b>RaiseEvent</b>  OpenParenthesis  ParameterList  CloseParenthesis  LineTerminator
    [  Block  ]
<b>End</b><b>RaiseEvent</b>  StatementTerminator</pre>

<pre>ConstantMemberDeclaration  ::=
    [  Attributes  ]  [  ConstantModifier+  ]  <b>Const</b>  ConstantDeclarators  StatementTerminator</pre>

<pre>ConstantModifier  ::=  AccessModifier  |  <b>Shadows</b></pre>

<pre>ConstantDeclarators  ::=
    ConstantDeclarator  |
    ConstantDeclarators  Comma  ConstantDeclarator</pre>

<pre>ConstantDeclarator  ::=  Identifier  [  <b>As</b>  TypeName  ]  Equals  ConstantExpression  StatementTerminator</pre>

<pre>VariableMemberDeclaration  ::=
    [  Attributes  ]  VariableModifier+  VariableDeclarators  StatementTerminator</pre>

<pre>VariableModifier  ::=
    AccessModifier  |
<b>Shadows</b>  |
<b>Shared</b>  |
<b>ReadOnly</b>  |
<b>WithEvents</b>  |
<b>Dim</b></pre>

<pre>VariableDeclarators  ::=
    VariableDeclarator  |
    VariableDeclarators  Comma  VariableDeclarator</pre>

<pre>VariableDeclarator  ::=
    VariableIdentifiers   <b>As</b>  ObjectCreationExpression  |
    VariableIdentifiers  [  <b>As</b>  TypeName  ]  [  Equals  Expression  ]</pre>

<pre>VariableIdentifiers  ::=
    VariableIdentifier  |
    VariableIdentifiers  Comma  VariableIdentifier</pre>

<pre>VariableIdentifier  ::=  Identifier  IdentifierModifiers</pre>

<pre>ArraySizeInitializationModifier  ::=
    OpenParenthesis  BoundList  CloseParenthesis  [  ArrayTypeModifiers  ]</pre>

<pre>BoundList::=
    Bound |
    BoundList  Comma  Bound</pre>

<pre>Bound  ::=
    Expression  |
<b>0</b><b>To</b>  Expression</pre>

<pre>PropertyMemberDeclaration  ::=
    RegularPropertyMemberDeclaration  |
    MustOverridePropertyMemberDeclaration  |
    AutoPropertyMemberDeclaration</pre>

<pre>PropertySignature  ::=
<b>Property</b>  Identifier  [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]
        [  <b>As</b>  [  Attributes  ]  TypeName  ]</pre>

<pre>RegularPropertyMemberDeclaration  ::=
    [  Attributes  ]  [  PropertyModifier+  ]  PropertySignature   [  ImplementsClause  ]  LineTerminator
    PropertyAccessorDeclaration+
<b>End</b><b>Property</b>  StatementTerminator</pre>

<pre>MustOverridePropertyMemberDeclaration  ::=
    [  Attributes  ]  MustOverridePropertyModifier+  PropertySignature  [  ImplementsClause  ]
        StatementTerminator</pre>

<pre>AutoPropertyMemberDeclaration  ::=
    [  Attributes  ]  [  AutoPropertyModifier+  ]  <b>Property</b>  Identifier
        [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]
        [  <b>As</b>  [  Attributes  ]  TypeName  ]  [  Equals  Expression  ]  [  ImplementsClause  ]  LineTerminator  |
    [  Attributes  ]  [  AutoPropertyModifier+  ]  <b>Property</b>  Identifier
        [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]
<b>As</b>  [  Attributes  ]  <b>New</b>  [  NonArrayTypeName
        [  OpenParenthesis  [  ArgumentList  ]  CloseParenthesis  ]  ]  [  ObjectCreationExpressionInitializer  ]
        [  ImplementsClause  ]  LineTerminator</pre>

<pre>InterfacePropertyMemberDeclaration  ::=
    [  Attributes  ]  [  InterfacePropertyModifier+  ]  PropertySignature  StatementTerminator</pre>

<pre>AutoPropertyModifier  ::=
    AccessModifier  |
<b>Shadows</b>  |
<b>Shared</b>  |
<b>Overridable</b>  |
<b>NotOverridable</b>  |
<b>Overrides</b>  |
<b>Overloads</b></pre>

<pre>PropertyModifier  ::=
    AutoPropertyModifier  |
<b>Default</b>  |
<b>ReadOnly</b>  |
<b>WriteOnly</b>  |
<b>Iterator</b></pre>

<pre>MustOverridePropertyModifier  ::=  PropertyModifier  |  <b>MustOverride</b></pre>

<pre>InterfacePropertyModifier  ::=
<b>Shadows</b>  |
<b>Overloads</b>  |
<b>Default</b>  |
<b>ReadOnly</b>  |
<b>WriteOnly</b></pre>

<pre>PropertyAccessorDeclaration  ::=  PropertyGetDeclaration  |  PropertySetDeclaration</pre>

<pre>PropertyGetDeclaration  ::=
    [  Attributes  ]  [  AccessModifier  ]  <b>Get</b>  LineTerminator
    [  Block  ]
<b>End</b><b>Get</b>  StatementTerminator</pre>

<pre>PropertySetDeclaration  ::=
    [  Attributes  ]  [  AccessModifier  ]  <b>Set</b>  [  OpenParenthesis  [  ParameterList  ]  CloseParenthesis  ]  LineTerminator
    [  Block  ]
<b>End</b><b>Set</b>  StatementTerminator</pre>

<pre>OperatorDeclaration  ::=
    [  Attributes  ]  [  OperatorModifier+  ]  <b>Operator</b>  OverloadableOperator  OpenParenthesis  ParameterList  CloseParenthesis
        [  <b>As</b>  [  Attributes  ]  TypeName  ]  LineTerminator
    [  Block  ]
<b>End</b><b>Operator</b>  StatementTerminator</pre>

<pre>OperatorModifier  ::=  <b>Public</b>  |  <b>Shared</b>  |  <b>Overloads</b>  |  <b>Shadows</b>  |  <b>Widening</b>  |  <b>Narrowing</b></pre>

<pre>OverloadableOperator  ::=
<b>+</b>  |  <b>-</b>  |  <b>*</b>  |  <b>/</b>  |  <b>\</b>  |  <b>&</b>  |  <b>Like</b>  |  <b>Mod</b>  |  <b>And</b>  |  <b>Or</b>  |  <b>Xor</b>  |  <b>^</b>  |  <b><</b><b><</b>  |  <b>></b><b>></b>  |
<b>=</b>  |  <b><</b><b>></b>  |  <b>></b>  |  <b><</b>  |  <b>></b><b>=</b>  |  <b><</b><b>=</b>  |  <b>Not</b>  |  <b>IsTrue</b>  |  <b>IsFalse</b>  |  <b>CType</b></pre>

### Statements

<pre>Statement  ::=
    LabelDeclarationStatement  |
    LocalDeclarationStatement  |
    WithStatement  |
    SyncLockStatement  |
    EventStatement  |
    AssignmentStatement  |
    InvocationStatement  |
    ConditionalStatement  |
    LoopStatement  |
    ErrorHandlingStatement  |
    BranchStatement  |
    ArrayHandlingStatement  |
    UsingStatement  |
    AwaitStatement  |
    YieldStatement</pre>

<pre>Block  ::=  [  Statements+  ]</pre>

<pre>LabelDeclarationStatement  ::=  LabelName  <b>:</b></pre>

<pre>LabelName  ::=  Identifier  |  IntLiteral</pre>

<pre>Statements  ::=
    [  Statement  ]  |
    Statements  <b>:</b>  [  Statement  ]</pre>

<pre>LocalDeclarationStatement  ::=  LocalModifier  VariableDeclarators  StatementTerminator</pre>

<pre>LocalModifier  ::=  <b>Static</b>  |  <b>Dim</b>  |  <b>Const</b></pre>

<pre>WithStatement  ::=
<b>With</b>  Expression  StatementTerminator
    [  Block  ]
<b>End</b><b>With</b>  StatementTerminator</pre>

<pre>SyncLockStatement  ::=
<b>SyncLock</b>  Expression  StatementTerminator
    [  Block  ]
<b>End</b><b>SyncLock</b>  StatementTerminator</pre>

<pre>EventStatement  ::=
    RaiseEventStatement  |
    AddHandlerStatement  |
    RemoveHandlerStatement</pre>

<pre>RaiseEventStatement  ::=  <b>RaiseEvent</b>  IdentifierOrKeyword  [  OpenParenthesis  [  ArgumentList  ]  CloseParenthesis  ]
    StatementTerminator</pre>

<pre>AddHandlerStatement  ::=  <b>AddHandler</b>  Expression  Comma  Expression  StatementTerminator</pre>

<pre>RemoveHandlerStatement  ::=  <b>RemoveHandler</b>  Expression  Comma  Expression  StatementTerminator</pre>

<pre>AssignmentStatement  ::=
    RegularAssignmentStatement  |
    CompoundAssignmentStatement  |
    MidAssignmentStatement</pre>

<pre>RegularAssignmentStatement  ::=  Expression  Equals  Expression  StatementTerminator</pre>

<pre>CompoundAssignmentStatement  ::=  Expression  CompoundBinaryOperator  [  LineTerminator  ]
        Expression  StatementTerminator</pre>

<pre>CompoundBinaryOperator  ::=  <b>^</b><b>=</b>  |  <b>*</b><b>=</b>  |  <b>/</b><b>=</b>  |  <b>\</b><b>=</b>  |  <b>+</b><b>=</b>  |  <b>-</b><b>=</b>  |  <b>&</b><b>=</b>  |  <b><</b><b><</b><b>=</b>  |  <b>></b><b>></b><b>=</b></pre>

<pre>MidAssignmentStatement  ::=
<b>Mid</b>  [  <b>$</b>  ]  OpenParenthesis  Expression  Comma  Expression  [  Comma  Expression  ]  CloseParenthesis
        Equals  Expression  StatementTerminator</pre>

<pre>InvocationStatement  ::=  [  <b>Call</b>  ]  InvocationExpression  StatementTerminator</pre>

<pre>ConditionalStatement  ::=  IfStatement  |  SelectStatement</pre>

<pre>IfStatement  ::=  BlockIfStatement  |  LineIfThenStatement</pre>

<pre>BlockIfStatement  ::=
<b>If</b>  BooleanExpression  [  <b>Then</b>  ]  StatementTerminator
    [  Block  ]
    [  ElseIfStatement+  ]
    [  ElseStatement  ]
<b>End</b><b>If</b>  StatementTerminator</pre>

<pre>ElseIfStatement  ::=
    ElseIf  BooleanExpression  [  <b>Then</b>  ]  StatementTerminator
    [  Block  ]</pre>

<pre>ElseStatement  ::=
<b>Else</b>  StatementTerminator
    [  Block  ]</pre>

<pre>LineIfThenStatement  ::=
<b>If</b>  BooleanExpression  <b>Then</b>  Statements  [  <b>Else</b>  Statements  ]  StatementTerminator</pre>

<pre>SelectStatement  ::=
<b>Select</b>  [  <b>Case</b>  ]  Expression  StatementTerminator
    [  CaseStatement+  ]
    [  CaseElseStatement  ]
<b>End</b><b>Select</b>  StatementTerminator</pre>

<pre>CaseStatement  ::=
<b>Case</b>  CaseClauses  StatementTerminator
    [  Block  ]</pre>

<pre>CaseClauses  ::=
    CaseClause  |
    CaseClauses  Comma  CaseClause</pre>

<pre>CaseClause  ::=
    [  <b>Is</b>  [  LineTerminator  ]  ]  ComparisonOperator  [ LineTerminator ] Expression  |
    Expression  [  <b>To</b>  Expression  ]</pre>

<pre>ComparisonOperator  ::=  <b>=</b>  |  <b><</b><b>></b>  |  <b><</b>  |  <b>></b>  |  <b>></b><b>=</b>  |  <b><</b><b>=</b></pre>

<pre>CaseElseStatement  ::=
<b>Case</b><b>Else</b>  StatementTerminator
    [  Block  ]</pre>

<pre>LoopStatement  ::=
    WhileStatement  |
    DoLoopStatement  |
    ForStatement  |
    ForEachStatement</pre>

<pre>WhileStatement  ::=
<b>While</b>  BooleanExpression  StatementTerminator
    [  Block  ]
<b>End</b><b>While</b>  StatementTerminator</pre>

<pre>DoLoopStatement  ::=  DoTopLoopStatement  |  DoBottomLoopStatement</pre>

<pre>DoTopLoopStatement  ::=
<b>Do</b>  [  WhileOrUntil  BooleanExpression  ]  StatementTerminator
    [  Block  ]
<b>Loop</b>  StatementTerminator</pre>

<pre>DoBottomLoopStatement  ::=
<b>Do</b>  StatementTerminator
    [  Block  ]
    Loop  WhileOrUntil  BooleanExpression  StatementTerminator</pre>

<pre>WhileOrUntil  ::=  <b>While</b>  |  <b>Until</b></pre>

<pre>ForStatement  ::=
<b>For</b>  LoopControlVariable  Equals  Expression  <b>To</b>  Expression  [  <b>Step</b>  Expression  ]  StatementTerminator
    [  Block  ]
    [  <b>Next</b>  [  NextExpressionList  ]  StatementTerminator  ]</pre>

<pre>LoopControlVariable  ::=
    Identifier  [  IdentifierModifiers  <b>As</b>  TypeName  ]  |
    Expression</pre>

<pre>NextExpressionList  ::=
    Expression  |
    NextExpressionList  Comma  Expression</pre>

<pre>ForEachStatement  ::=
<b>For</b><b>Each</b>  LoopControlVariable  <b>In</b>  [  LineTerminator  ]  Expression  StatementTerminator
    [  Block  ]
    [  <b>Next</b>  [  NextExpressionList  ]  StatementTerminator  ]</pre>

<pre>ErrorHandlingStatement  ::=
    StructuredErrorStatement  |
    UnstructuredErrorStatement</pre>

<pre>StructuredErrorStatement  ::=
    ThrowStatement  |
    TryStatement</pre>

<pre>TryStatement  ::=
<b>Try</b>  StatementTerminator
    [  Block  ]
    [  CatchStatement+  ]
    [  FinallyStatement  ]
<b>End</b><b>Try</b>  StatementTerminator</pre>

<pre>FinallyStatement  ::=
<b>Finally</b>  StatementTerminator
    [  Block  ]</pre>

<pre>CatchStatement  ::=
<b>Catch</b>  [  Identifier  [  <b>As</b>  NonArrayTypeName  ]  ]  [  <b>When</b>  BooleanExpression  ]  StatementTerminator
    [  Block  ]</pre>

<pre>ThrowStatement  ::=  <b>Throw</b>  [  Expression  ]  StatementTerminator</pre>

<pre>UnstructuredErrorStatement  ::=
    ErrorStatement  |
    OnErrorStatement  |
    ResumeStatement</pre>

<pre>ErrorStatement  ::=  <b>Error</b>  Expression  StatementTerminator</pre>

<pre>OnErrorStatement  ::=  <b>On</b><b>Error</b>  ErrorClause  StatementTerminator</pre>

<pre>ErrorClause  ::=
<b>GoTo</b><b>-</b><b>1</b>  |
<b>GoTo</b><b>0</b>  |
    GoToStatement  |
<b>Resume</b><b>Next</b></pre>

<pre>ResumeStatement  ::=  <b>Resume</b>  [  ResumeClause  ]  StatementTerminator</pre>

<pre>ResumeClause  ::=  <b>Next</b>  |  LabelName</pre>

<pre>BranchStatement  ::=
    GoToStatement  |
    ExitStatement  |
    ContinueStatement  |
    StopStatement  |
    EndStatement  |
    ReturnStatement</pre>

<pre>GoToStatement  ::=  <b>GoTo</b>  LabelName  StatementTerminator</pre>

<pre>ExitStatement  ::=  <b>Exit</b>  ExitKind  StatementTerminator</pre>

<pre>ExitKind  ::=  <b>Do</b>  |  <b>For</b>  |  <b>While</b>  |  <b>Select</b>  |  <b>Sub</b>  |  <b>Function</b>  |  <b>Property</b>  |  <b>Try</b></pre>

<pre>ContinueStatement  ::=  <b>Continue</b>  ContinueKind  StatementTerminator</pre>

<pre>ContinueKind  ::=  <b>Do</b>  |  <b>For</b>  |  <b>While</b></pre>

<pre>StopStatement  ::=  <b>Stop</b>  StatementTerminator</pre>

<pre>EndStatement  ::=  <b>End</b>  StatementTerminator</pre>

<pre>ReturnStatement  ::=  <b>Return</b>  [  Expression  ]  StatementTerminator</pre>

<pre>ArrayHandlingStatement  ::=
    RedimStatement  |
    EraseStatement</pre>

<pre>RedimStatement  ::=  <b>ReDim</b>  [  <b>Preserve</b>  ]  RedimClauses  StatementTerminator</pre>

<pre>RedimClauses  ::=
    RedimClause  |
    RedimClauses  Comma  RedimClause</pre>

<pre>RedimClause  ::=  Expression  ArraySizeInitializationModifier</pre>

<pre>EraseStatement  ::=  <b>Erase</b>  EraseExpressions  StatementTerminator</pre>

<pre>EraseExpressions  ::=
    Expression  |
    EraseExpressions  Comma  Expression</pre>

<pre>UsingStatement  ::=
<b>Using</b>  UsingResources  StatementTerminator
    [  Block  ]
<b>End</b><b>Using</b>  StatementTerminator</pre>

<pre>UsingResources  ::=  VariableDeclarators  |  Expression</pre>

<pre>AwaitStatement  ::=  AwaitOperatorExpression  StatementTerminator</pre>

<pre>YieldStatement  ::=  <b>Yield</b>  Expressions  StatementTerminator</pre>

### Expressions

<pre>Expression  ::=
    SimpleExpression  |
    TypeExpression  |
    MemberAccessExpression  |
    DictionaryAccessExpression  |
*InvocationExpression* |
    IndexExpression  |
    NewExpression  |
    CastExpression  |
    OperatorExpression  |
    ConditionalExpression  |
    LambdaExpression  |
    QueryExpression  |
    XMLLiteralExpression  |
    XMLMemberAccessExpression</pre>

<pre>ConstantExpression  ::=  Expression</pre>

<pre>SimpleExpression  ::=
    LiteralExpression  |
    ParenthesizedExpression  |
    InstanceExpression  |
    SimpleNameExpression  |
    AddressOfExpression</pre>

<pre>LiteralExpression  ::=  Literal</pre>

<pre>ParenthesizedExpression  ::=  OpenParenthesis  Expression  CloseParenthesis</pre>

<pre>InstanceExpression  ::=  <b>Me</b></pre>

<pre>SimpleNameExpression  ::=  Identifier  [  OpenParenthesis  <b>Of</b>  TypeArgumentList  CloseParenthesis  ]</pre>

<pre>AddressOfExpression  ::=  <b>AddressOf</b>  Expression</pre>

<pre>TypeExpression  ::=
    GetTypeExpression  |
    TypeOfIsExpression  |
    IsExpression  |
    GetXmlNamespaceExpression</pre>

<pre>GetTypeExpression  ::=  <b>GetType</b>  OpenParenthesis  GetTypeTypeName  CloseParenthesis</pre>

<pre>GetTypeTypeName  ::=
    TypeName  |
*Qualified*OpenTypeName</pre>

<pre>QualifiedOpenTypeName  ::=
    Identifier  [  TypeArityList  ]  |
<b>Global</b>  Period  IdentifierOrKeyword    [  TypeArityList  ]  |
    QualifiedOpenTypeName  Period  IdentifierOrKeyword  [  TypeArityList  ]</pre>

<pre>*TypeArityList*  ::=  `(``Of`  [ *CommaList* ] `)`</pre>

<pre>CommaList  ::=
    Comma  |
    CommaList  Comma</pre>

<pre>TypeOfIsExpression  ::=  <b>TypeOf</b>  Expression  <b>Is</b>  [  LineTerminator  ]  TypeName</pre>

<pre>IsExpression  ::=
    Expression  <b>Is</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>IsNot</b>  [  LineTerminator  ]  Expression</pre>

<pre>GetXmlNamespaceExpression  ::=  <b>Get</b><b>XmlNamespace</b>  OpenParenthesis  [  XMLNamespaceName  ]
        CloseParenthesis</pre>

<pre>MemberAccessExpression  ::=
     [  MemberAccessBase  ]  *Period*  IdentifierOrKeyword
        [  OpenParenthesis  <b>Of</b>  TypeArgumentList  CloseParenthesis  ]</pre>

<pre>MemberAccessBase  ::=
    Expression  |
    BuiltInTypeName  |
<b>Global</b>  |
<b>MyClass</b>  |
<b>MyBase</b></pre>

<pre>DictionaryAccessExpression  ::=  [  Expression  ]  <b>!</b>  IdentifierOrKeyword</pre>

<pre>InvocationExpression  ::=  Expression  [  OpenParenthesis  [  ArgumentList  ]  CloseParenthesis  ]</pre>

<pre>ArgumentList  ::=    PositionalArgumentList  |
    PositionalArgumentList  Comma  NamedArgumentList  |
    NamedArgumentList</pre>

<pre>PositionalArgumentList  ::=
    [ Expression  ] |
    PositionalArgumentList  Comma  [  Expression  ]</pre>

<pre>NamedArgumentList  ::=
    IdentifierOrKeyword  ColonEquals  Expression  |
    NamedArgumentList  Comma  IdentifierOrKeyword  ColonEquals  Expression</pre>

<pre>IndexExpression  ::=  Expression  OpenParenthesis  [  ArgumentList  ]  CloseParenthesis</pre>

<pre>NewExpression  ::=
    ObjectCreationExpression  |
    ArrayExpression  |
    AnonymousObjectCreationExpression</pre>

<pre>AnonymousObjectCreationExpression  ::=
<b>New</b>  ObjectMemberInitializer</pre>

<pre>ObjectCreationExpression  ::=
<b>New</b>  NonArrayTypeName  [  OpenParenthesis  [  ArgumentList  ]  CloseParenthesis  ]
        [  ObjectCreationExpressionInitializer  ]</pre>

<pre>ObjectCreationExpressionInitializer  ::=  ObjectMemberInitializer  |  ObjectCollectionInitializer</pre>

<pre>ObjectMemberInitializer  ::=
<b>With</b>  OpenCurlyBrace  FieldInitializerList  CloseCurlyBrace</pre>

<pre>FieldInitializerList  ::=
    FieldInitializer  |
    FieldInitializerList  `,`  FieldInitializer</pre>

<pre>FieldInitializer  ::=  [  [  <b>Key</b>  ]  `.`  IdentifierOrKeyword  `=`  ]  Expression</pre>

<pre>ObjectCollectionInitializer  ::=  <b>From</b>  CollectionInitializer</pre>

<pre>CollectionInitializer  ::=  OpenCurlyBrace  [  CollectionElementList  ]  CloseCurlyBrace</pre>

<pre>CollectionElementList  ::=
    CollectionElement  |
    CollectionElementList  Comma  CollectionElement</pre>

<pre>CollectionElement  *::=*
*Expression*  |
    CollectionInitializer</pre>

<pre>ArrayExpression * ::=*
*ArrayCreationExpression*  |
*ArrayLiteralExpression*</pre>

<pre>ArrayCreationExpression  ::=
<b>New</b>  NonArrayTypeName  ArrayNameModifier  CollectionInitializerArrayLiteralExpression  ::=  CollectionInitializer</pre>

<pre>CastExpression  ::=
<b>DirectCast</b>  OpenParenthesis  Expression  Comma  TypeName  CloseParenthesis  |
<b>TryCast</b>  OpenParenthesis  Expression  Comma  TypeName  CloseParenthesis  |
<b>CType</b>  OpenParenthesis  Expression  Comma  TypeName  CloseParenthesis  |
    CastTarget  OpenParenthesis  Expression  CloseParenthesis</pre>

<pre>CastTarget  ::=
<b>CBool</b>  |  <b>CByte</b>  |  <b>CChar</b>  |  <b>CDate</b>  |  <b>CDec</b>  |  <b>CDbl</b>  |  <b>CInt</b>  |  <b>CLng</b>  |  <b>CObj</b>  |  <b>CSByte</b>  |  <b>CShort</b>  |
<b>CSng</b>  |    <b>CStr</b>  |  <b>CUInt</b>  |  <b>CULng</b>  |  <b>CUShort</b></pre>

<pre>OperatorExpression  ::=
    ArithmeticOperatorExpression  |
    RelationalOperatorExpression  |
    LikeOperatorExpression  |
    ConcatenationOperatorExpression  |
    ShortCircuitLogicalOperatorExpression  |
    LogicalOperatorExpression  |
    ShiftOperatorExpression  |
    AwaitOperatorExpression</pre>

<pre>ArithmeticOperatorExpression  ::=
    UnaryPlusExpression  |
    UnaryMinusExpression  |
    AdditionOperatorExpression  |
    SubtractionOperatorExpression  |
    MultiplicationOperatorExpression  |
    DivisionOperatorExpression  |
    ModuloOperatorExpression  |
    ExponentOperatorExpression</pre>

<pre>UnaryPlusExpression  ::=  <b>+</b>  Expression</pre>

<pre>UnaryMinusExpression  ::=  <b>-</b>  Expression</pre>

<pre>AdditionOperatorExpression  ::=  Expression  <b>+</b>  [  LineTerminator  ]  Expression</pre>

<pre>SubtractionOperatorExpression  ::=  Expression  <b>-</b>  [  LineTerminator  ]  Expression</pre>

<pre>MultiplicationOperatorExpression  ::=  Expression  <b>*</b>  [  LineTerminator  ]  Expression</pre>

<pre>DivisionOperatorExpression  ::=
    FPDivisionOperatorExpression  |
    IntegerDivisionOperatorExpression</pre>

<pre>FPDivisionOperatorExpression  ::=  Expression  <b>/</b>  [  LineTerminator  ]  Expression</pre>

<pre>IntegerDivisionOperatorExpression  ::=  Expression  <b>\</b>  [  LineTerminator  ]  Expression</pre>

<pre>ModuloOperatorExpression  ::=  Expression  <b>Mod</b>  [  LineTerminator  ]  Expression</pre>

<pre>ExponentOperatorExpression  ::=  Expression  <b>^</b>  [  LineTerminator  ]  Expression</pre>

<pre>RelationalOperatorExpression  ::=
    Expression  <b>=</b>  [  LineTerminator  ]  Expression  |
    Expression  <b><</b><b>></b>  [  LineTerminator  ]  Expression  |
    Expression  <b><</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>></b>  [  LineTerminator  ]  Expression  |
    Expression  <b><</b><b>=</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>></b><b>=</b>  [  LineTerminator  ]  Expression</pre>

<pre>LikeOperatorExpression  ::=  Expression  <b>Like</b>  [  LineTerminator  ]  Expression</pre>

<pre>ConcatenationOperatorExpression  ::=  Expression  <b>&</b>  [  LineTerminator  ]  Expression</pre>

<pre>LogicalOperatorExpression  ::=
<b>Not</b>  Expression  |
    Expression  <b>And</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>Or</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>Xor</b>  [  LineTerminator  ]  Expression</pre>

<pre>ShortCircuitLogicalOperatorExpression  ::=
    Expression  <b>AndAlso</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>OrElse</b>  [  LineTerminator  ]  Expression</pre>

<pre>ShiftOperatorExpression  ::=
    Expression  <b><</b><b><</b>  [  LineTerminator  ]  Expression  |
    Expression  <b>></b><b>></b>  [  LineTerminator  ]  Expression</pre>

<pre>BooleanExpression  ::=  Expression</pre>

<pre>LambdaExpression  ::=
    SingleLineLambda  |
    MultiLineLambda</pre>

<pre>SingleLineLambda  ::=
    [  LambdaModifier+  ]  <b>Function</b>  [  OpenParenthesis  [  ParametertList  ]  CloseParenthesis  ]  Expression  |
    [  LambdaModifier+  ]  <b>Sub</b>  [  OpenParenthesis  [  ParametertList  ]  CloseParenthesis  ]  Statement</pre>

<pre>MultiLineLambda  ::=
    MultiLineFunctionLambda  |
    MultiLineSubLambda</pre>

<pre>MultiLineFunctionLambda  ::=
    [  LambdaModifier+  ]  <b>Function</b>  [  OpenParenthesis  [  ParametertList  ]  CloseParenthesis  ]  [  <b>As</b>  TypeName  ]  LineTerminator
    Block
<b>End</b><b>Function</b></pre>

<pre>MultiLineSubLambda  ::=
    [  LambdaModifier+  ]  <b>Sub</b>  [  OpenParenthesis  [  ParametertList  ]  CloseParenthesis  ]  LineTerminator
    Block
<b>End</b><b>Sub</b></pre>

<pre>LambdaModifier  ::=
<b>Async</b>  |
<b>Iterator</b></pre>

<pre>QueryExpression  ::=  
    FromOrAggregateQueryOperator  |
    QueryExpression  QueryOperator</pre>

<pre>FromOrAggregateQueryOperator  ::=  FromQueryOperator  |  AggregateQueryOperator</pre>

<pre>JoinOrGroupJoinQueryOperator * := * JoinQueryOperator * | * GroupJoinQueryOperator</pre>

<pre>QueryOperator ::=
    FromQueryOperator  |
    AggregateQueryOperator  |
    SelectQueryOperator  |
    DistinctQueryOperator  |
    WhereQueryOperator  |
    OrderByQueryOperator  |
    PartitionQueryOperator  |
    LetQueryOperator |
    GroupByQueryOperator  |    *JoinOr*GroupJoinQueryOperator</pre>

<pre>CollectionRangeVariableDeclarationList ::=
    CollectionRangeVariableDeclaration  |
    CollectionRangeVariableDeclarationList  Comma  CollectionRangeVariableDeclaration</pre>

<pre>CollectionRangeVariableDeclaration ::=  
    Identifier  [  <b>As</b>  TypeName  ]  <b>In</b>  [  LineTerminator  ]  Expression</pre>

<pre>ExpressionRangeVariableDeclarationList ::=
    ExpressionRangeVariableDeclaration  |
    ExpressionRangeVariableDeclarationList  Comma  ExpressionRangeVariableDeclaration</pre>

<pre>ExpressionRangeVariableDeclaration ::=  
    Identifier  [  <b>As</b>  TypeName  ]  Equals  Expression</pre>

<pre>FromQueryOperator ::=
    [  LineTerminator  ]  <b>From</b>  [  LineTerminator  ]  CollectionRangeVariableDeclarationList</pre>

<pre>JoinQueryOperator  ::=
    [  LineTerminator  ]  <b>Join</b>  [  LineTerminator  ]  CollectionRangeVariableDeclaration
        [  JoinOrGroupJoinQueryOperator  ]  [  LineTerminator  ]  <b>On</b>  [  LineTerminator  ]  JoinConditionList</pre>

<pre>JoinConditionList  ::=
    JoinCondition  |
    JoinConditionList  <b>And</b>  [  LineTerminator  ]  JoinCondition</pre>

<pre>JoinCondition  ::=  Expression  <b>Equals</b>  [  LineTerminator  ]  Expression</pre>

<pre>LetQueryOperator ::=
    [  LineTerminator  ]  <b>Let</b>  [  LineTerminator  ]  ExpressionRangeVariableDeclarationList</pre>

<pre>SelectQueryOperator  ::=
    [  LineTerminator  ]  <b>Select</b>  [  LineTerminator  ]  ExpressionRangeVariableDeclarationList</pre>

<pre>DistinctQueryOperator ::=
    [  LineTerminator  ]  <b>Distinct</b>  [  LineTerminator  ]  </pre>

<pre>WhereQueryOperator ::=  
    [  LineTerminator  ]  <b>Where</b>  [  LineTerminator  ]  BooleanExpression</pre>

<pre>PartitionQueryOperator ::=  
    [  LineTerminator  ]  <b>Take</b>  [  LineTerminator  ]  Expression  |
      [  LineTerminator  ]  <b>Take</b><b>While</b>  [  LineTerminator  ]  BooleanExpression  |
    [  LineTerminator  ]  <b>Skip</b>  [  LineTerminator  ]  Expression  |
    [  LineTerminator  ]  <b>Skip</b><b>While</b>  [  LineTerminator  ]  BooleanExpression  </pre>

<pre>OrderByQueryOperator  ::=
    [  LineTerminator  ]  <b>Order</b><b>By</b>  [  LineTerminator  ]  OrderExpressionList</pre>

<pre>OrderExpressionList  ::=
    OrderExpression  |
    OrderExpressionList  Comma  OrderExpression</pre>

<pre>OrderExpression  ::=
    Expression  [  Ordering  ]</pre>

<pre>Ordering  ::=  <b>Ascending</b>  |  <b>Descending</b></pre>

<pre>GroupByQueryOperator  ::=
    [  LineTerminator  ]  <b>Group</b>  [ [  LineTerminator  ]  ExpressionRangeVariableDeclarationList ]
        [  LineTerminator  ]  <b>By</b>  [  LineTerminator  ]  ExpressionRangeVariableDeclarationList
        [  LineTerminator  ]  <b>Into</b>  [  LineTerminator  ]  ExpressionRangeVariableDeclarationList</pre>

<pre>AggregateQueryOperator ::=
    [  LineTerminator  ]  <b>Aggregate</b>  [  LineTerminator  ]  CollectionRangeVariableDeclaration
        [  QueryOperator+  ]
        [  LineTerminator  ]  <b>Into</b>  [  LineTerminator  ]  ExpressionRangeVariableDeclarationList</pre>

<pre>GroupJoinQueryOperator  ::=
    [  LineTerminator  ]  <b>Group</b><b>Join</b>  [  LineTerminator  ]  CollectionRangeVariableDeclaration
         [  JoinOrGroupJoinQueryOperator  ]   [  LineTerminator  ]  <b>On</b>  [  LineTerminator  ]  JoinConditionList
        [  LineTerminator  ]  <b>Into</b>  [  LineTerminator  ]  ExpressionRangeVariableDeclarationList</pre>

<pre>ConditionalExpression  ::=  
<b>If</b>  OpenParenthesis  BooleanExpression  Comma  Expression  Comma  Expression  CloseParenthesis  |
<b>If</b>  OpenParenthesis  Expression  Comma  Expression  CloseParenthesis</pre>

<pre>XMLLiteralExpression  ::=
    XMLDocument  |
    XMLElement  |    XMLProcessingInstruction  |
    XMLComment  |
    XMLCDATASection</pre>

<pre>XMLCharacter  ::=
    < Unicode tab character (0x0009) >  |
    < Unicode linefeed character (0x000A) >  |
    < Unicode carriage return character (0x000D) >  |
    < Unicode characters 0x0020 – 0xD7FF >  |
    < Unicode characters 0xE000 – 0xFFFD >  |
    < Unicode characters 0x10000 – 0x10FFFF ></pre>

<pre>XMLString  ::=  XMLCharacter+</pre>

<pre>XMLWhitespace  ::=  XMLWhitespaceCharacter+</pre>

<pre>XMLWhitespaceCharacter  ::=
    < Unicode carriage return character (0x000D) >  |
    < Unicode linefeed character (0x000A) >  |
    < Unicode space character (0x0020) >  |
    < Unicode tab character (0x0009) ></pre>

<pre>XMLNameCharacter  ::= XMLLetter  |  XMLDigit  |  <b>.</b>  |  <b>-</b>  |  <b>_</b>  |  <b>:</b>  |  XMLCombiningCharacter  |  XMLExtender </pre>

<pre>XMLNameStartCharacter  ::=  XMLLetter  |  _  |  :</pre>

<pre>XMLName  ::=  XMLNameStartCharacter  [  XMLNameCharacter+  ] </pre>

<pre>XMLLetter  ::=  
    < Unicode character as defined in the Letter production of the XML 1.0 specification ></pre>

<pre>XMLDigit  ::=
    < Unicode character as defined in the Digit production of the XML 1.0 specification ></pre>

<pre>XMLCombiningCharacter  ::=
    < Unicode character as defined in the CombiningChar production of the XML 1.0 specification ></pre>

<pre>XMLExtender  ::=
    < Unicode character as defined in the Extender production of the XML 1.0 specification ></pre>

<pre>XMLEmbeddedExpression  ::=
<b><</b><b>%</b><b>=</b>  [  LineTerminator  ]  Expression  [  LineTerminator  ]  <b>%</b><b>></b></pre>

<pre>XMLDocument  ::=
    XMLDocumentPrologue  [  XMLMisc+  ]  XMLDocumentBody  [  XMLMisc+  ]</pre>

<pre>XMLDocumentPrologue  ::=
<b><</b><b>?</b><b>xml</b>  XMLVersion  [  XMLEncoding  ]  [  XMLStandalone  ]  [  XMLWhitespace  ]  <b>?</b><b>></b></pre>

<pre>XMLVersion  ::=
    XMLWhitespace  <b>version</b>  [  XMLWhitespace  ]  <b>=</b>  [  XMLWhitespace  ]  XMLVersionNumberValue</pre>

<pre>XMLVersionNumberValue  ::=  
    SingleQuoteCharacter  <b>1</b><b>.</b><b>0</b>  SingleQuoteCharacter  |
    DoubleQuoteCharacter  <b>1</b><b>.</b><b>0</b>  DoubleQuoteCharacter</pre>

<pre>XMLEncoding  ::=
    XMLWhitespace  <b>encoding</b>  [  XMLWhitespace  ]  <b>=</b>  [  XMLWhitespace  ]  XMLEncodingNameValue</pre>

<pre>XMLEncodingNameValue  ::=  
    SingleQuoteCharacter  XMLEncodingName  SingleQuoteCharacter  |
    DoubleQuoteCharacter  XMLEncodingName  DoubleQuoteCharacter</pre>

<pre>XMLEncodingName  ::=  XMLLatinAlphaCharacter  [  XMLEncodingNameCharacter+  ]</pre>

<pre>XMLEncodingNameCharacter  ::=
    XMLUnderscoreCharacter  |
    XMLLatinAlphaCharacter  |
    XMLNumericCharacter  |
    XMLPeriodCharacter  |
    XMLDashCharacter</pre>

<pre>XMLLatinAlphaCharacter  ::=
    < Unicode Latin alphabetic character (0x0041-0x005a, 0x0061-0x007a) ></pre>

<pre>XMLNumericCharacter  ::=  < Unicode digit character (0x0030-0x0039) ></pre>

<pre>XMLHexNumericCharacter  ::=
    XMLNumericCharacter  |
    < Unicode Latin hex alphabetic character (0x0041-0x0046, 0x0061-0x0066) ></pre>

<pre>XMLPeriodCharacter  ::=  < Unicode period character (0x002e) ></pre>

<pre>XMLUnderscoreCharacter  ::=  < Unicode underscore character (0x005f) ></pre>

<pre>XMLDashCharacter  ::=  < Unicode dash character (0x002d) ></pre>

<pre>XMLStandalone  ::=
    XMLWhitespace  <b>standalone</b>  [  XMLWhitespace  ]  <b>=</b>  [  XMLWhitespace  ]  XMLYesNoValue</pre>

<pre>XMLYesNoValue  ::=  
    SingleQuoteCharacter  XMLYesNo  SingleQuoteCharacter  |
    DoubleQuoteCharacter  XMLYesNo  DoubleQuoteCharacter</pre>

<pre>XMLYesNo  ::=  <b>yes</b>  |  <b>no</b></pre>

<pre>XMLMisc  ::=
    XMLComment  |
    XMLProcessingInstruction  |
    XMLWhitespace</pre>

<pre>XMLDocumentBody  ::=  XMLElement  |  XMLEmbeddedExpression</pre>

<pre>XMLElement  ::=
    XMLEmptyElement  |
    XMLElementStart  XMLContent  XMLElementEnd</pre>

<pre>XMLEmptyElement  ::=
<b><</b>  XMLQualifiedNameOrExpression  [  XMLAttribute+  ]  [  XMLWhitepace  ]  <b>/</b><b>></b></pre>

<pre>XMLElementStart  ::=
<b><</b>  XMLQualifiedNameOrExpression  [  XMLAttribute+  ]  [  XMLWhitespace  ]  <b>></b></pre>

<pre>XMLElementEnd  ::=
<b><</b><b>/</b><b>></b>  |
<b><</b><b>/</b>  XMLQualifiedName  [  XMLWhitespace  ]  <b>></b></pre>

<pre>XMLContent  ::=
    [  XMLCharacterData  ]  [  XMLNestedContent  [  XMLCharacterData  ]  ]+</pre>

<pre>XMLCharacterData  ::=
    < Any XMLCharacterDataString that does not contain the string "`]]>`" ></pre>

<pre>XMLCharacterDataString  ::=
    < Any Unicode character except `<` or `&` >+</pre>

<pre>XMLNestedContent  ::=
    XMLElement  |
    XMLReference  |
    XMLCDATASection  |
    XMLProcessingInstruction  |
    XMLComment  |
    XMLEmbeddedExpression</pre>

<pre>XMLAttribute  ::=
    XMLWhitespace  XMLAttributeName  [  XMLWhitespace  ]  <b>=</b>  [  XMLWhitespace  ]  XMLAttributeValue  |
    XMLWhitespace  XMLEmbeddedExpression</pre>

<pre>XMLAttributeName  ::=
    XMLQualifiedNameOrExpression  |
    XMLNamespaceAttributeName</pre>

<pre>XMLAttributeValue  ::=
    DoubleQuoteCharacter  [  XMLAttributeDoubleQuoteValueCharacter+  ]  DoubleQuoteCharacter  |
    SingleQuoteCharacter  [  XMLAttributeSingleQuoteValueCharacter+  ]  SingleQuoteCharacter  |
    XMLEmbeddedExpression</pre>

<pre>XMLAttributeDoubleQuoteValueCharacter  ::= 
    < Any XMLCharacter except `<`, `&`, or DoubleQuoteCharacter >  |
    XMLReference</pre>

<pre>XMLAttributeSingleQuoteValueCharacter ::=
    < Any XMLCharacter except `<`, `&`, or SingleQuoteCharacter >  |
    XMLReference</pre>

<pre>XMLReference  ::=  XMLEntityReference  |  XMLCharacterReference</pre>

<pre>XMLEntityReference  ::=
<b>&</b>  XMLEntityName  ;</pre>

<pre>XMLEntityName  ::=  <b>lt</b>  |  <b>gt</b>  |  <b>amp</b>  |  <b>apos</b>  |  <b>quot</b></pre>

<pre>XMLCharacterReference  ::=
<b>&</b><b>#</b>  XMLNumericCharacter+  <b>;</b>  |
<b>&</b><b>#</b><b>x</b>  XMLHexNumericCharacter+  <b>;</b></pre>

<pre>XMLNamespaceAttributeName ::=
    XMLPrefixedNamespaceAttributeName  |
    XMLDefaultNamespaceAttributeName</pre>

<pre>XMLPrefixedNamespaceAttributeName  ::=
<b>xmlns</b><b>:</b>  XMLNamespaceName</pre>

<pre>XMLDefaultNamespaceAttributeName  ::=
<b>xmlns</b></pre>

<pre>XMLNamespaceName  ::=  XMLNamespaceNameStartCharacter  [  XMLNamespaceNameCharacter+  ]</pre>

<pre>XMLNamespaceNameStartCharacter  ::=
    < Any XMLNameCharacter except `:`  ></pre>

<pre>XMLNamespaceNameCharacter  ::=  XMLLetter  |  _</pre>

<pre>XMLQualifiedNameOrExpression  ::=  XMLQualifiedName  |  XMLEmbeddedExpression</pre>

<pre>XMLQualifiedName  ::=
    XMLPrefixedName  |
    XMLUnprefixedName</pre>

<pre>XMLPrefixedName  ::=  XMLNamespaceName  :  XMLNamespaceName</pre>

<pre>XMLUnprefixedName  ::=  XMLNamespaceName</pre>

<pre>XMLProcessingInstruction  ::=
<b><</b><b>?</b>  XMLProcessingTarget  [  XMLWhitespace  [  XMLProcessingValue  ]  ]  <b>?</b><b>></b></pre>

<pre>XMLProcessingTarget  ::=
    < Any XMLName except a casing permutation of the string "`xml`" ></pre>

<pre>XMLProcessingValue  ::=
    < Any XMLString that does not contain the string "`?>``"` ></pre>

<pre>XMLComment  ::=
<b><</b><b>!</b><b>-</b><b>-</b>  [  XMLCommentCharacter+  ]  <b>-</b><b>-</b><b>></b></pre>

<pre>XMLCommentCharacter  ::=
    < Any XMLCharacter except dash (0x002D) >  |
<b>-</b>  < Any XMLCharacter except dash (0x002D) ></pre>

<pre>XMLCDATASection  ::=
<b><</b><b>!</b><b>[</b><b>CDATA</b><b>[</b>  [  XMLCDATASectionString  ]  <b>]</b><b>]</b><b>></b></pre>

<pre>XMLCDATASectionString  ::=
    < Any XMLString that does not contain the string "`]]>`" ></pre>

<pre>XMLMemberAccessExpression  ::=
    Expression  <b>.</b>  [  LineTerminator  ]  <b><</b>  XMLQualifiedName  <b>></b>  |
    Expression  <b>.</b>  [  LineTerminator  ]  <b>@</b>  [  LineTerminator  ]  <b><</b>  XMLQualifiedName  <b>></b>  |
    Expression  <b>.</b>  [  LineTerminator  ]  <b>@</b>  [  LineTerminator  ]  IdentifierOrKeyword  |
    Expression  <b>.</b><b>.</b><b>.</b>  [  LineTerminator  ]  <b><</b>  XMLQualifiedName  <b>></b></pre>

<pre>AwaitOperatorExpression  ::=  <b>Await</b>  Expression</pre>
