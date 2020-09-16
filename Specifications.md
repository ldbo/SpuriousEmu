The VBA language specifications can be found in MS-VBAL. It defines five levels of grammar:
    - the physical line grammar indicates how to split a file into lines
    - the logical line grammar indicates how to extract individual processable lines, grouping split statements split statements and splitting statements sharing a single line
    - lexical tokens are i

# Physical Line Grammar

Fully supported

# Logical Line Grammar

 - WSC: missing most-Unicode-class-Zs

# Lexical Tokens Grammar

 - Unicode characters should not be supported, see (here)[https://lark-parser.readthedocs.io/en/latest/classes.html#using-unicode-character-classes-with-regex]
 - Comment: discard line continuation since it seems to be a specification error
 - Homemade STRING
 - Only handle Latin identifier in LEXICAL_IDENTIFIER
 - IDENTIFIER does not exclude RESERVED_IDENTIFIER, forbidden use of identifiers is checked by the Transformer
 - Is currently case-sensitive, keywords using camel case

# Conditional Compilation grammar

Not supported

# Module Structure
 
 - Module header is not optional, so simple scripts can't be parsed as of now.

# Module Code section

 - statement-label-definition can occur elsewhere than at LINE-START
 - statement-label-definition not supported at sub/function/property footer
 - property-parameters is replaced by procedure-parameters in property-lhs-header, arity and arguments name are checked during compilation
 - 5.4.2 / 5.4.3 / 5.4.5 TBD
 - Expressions:
     // + value-expression includes l-expression to solve precedence issues
     + cc-expression not supported
     + constrained expressions need to have static semantic
     + member_access_expression includes simple_name_expression
