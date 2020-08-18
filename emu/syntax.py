"""Define the grammar of a VBA source file and implement a syntactic parser."""

from typing import List, Union

from pyparsing import (Forward, ParseException, ParserElement, QuotedString,
                       Regex, Suppress, StringEnd, StringStart, Word,
                       delimitedList, infixNotation, nums, Or, Keyword,
                       FollowedBy)
from pyparsing import Optional as pOptional

from .abstract_syntax_tree import *
from .error import ParsingError
from .operator import *
from .partial_block import *
from .preprocessor import Instruction, Preprocessor
from .type import Type, types

############
#  Tokens  #
############

# SEPARATOR

lparen = Suppress('(').setName('lparen')
rparen = Suppress(')').setName('rparen')
dot = Suppress('.').setName('dot')

# STATEMENT KEYWORD

call_kw = Suppress(Keyword('Call')).setName('call')
case_kw = Suppress(Keyword('Case')).setName('case')
close_kw = Suppress(Keyword('Close')).setName('close')
const_kw = Suppress(Keyword('Const')).setName('const')
declare_kw = Suppress(Keyword('Declare')).setName('declare')
defbool_kw = Suppress(Keyword('DefBool')).setName('defbool')
defbyte_kw = Suppress(Keyword('DefByte')).setName('defbyte')
defcur_kw = Suppress(Keyword('DefCur')).setName('defcur')
defdate_kw = Suppress(Keyword('DefDate')).setName('defdate')
defdbl_kw = Suppress(Keyword('DefDbl')).setName('defdbl')
defint_kw = Suppress(Keyword('DefInt')).setName('defint')
deflng_kw = Suppress(Keyword('DefLng')).setName('deflng')
deflnglng_kw = Suppress(Keyword('DefLngLng')).setName('deflnglng')
deflngptr_kw = Suppress(Keyword('DefLngPtr')).setName('deflngptr')
defobj_kw = Suppress(Keyword('DefObj')).setName('defobj')
defsng_kw = Suppress(Keyword('DefSng')).setName('defsng')
defstr_kw = Suppress(Keyword('DefStr')).setName('defstr')
defvar_kw = Suppress(Keyword('DefVar')).setName('defvar')
dim_kw = Suppress(Keyword('Dim')).setName('dim')
do_kw = Suppress(Keyword('Do')).setName('do')
else_kw = Suppress(Keyword('Else')).setName('else')
elseif_kw = Suppress(Keyword('ElseIf')).setName('elseif')
end_kw = Suppress(Keyword('End')).setName('end')
endif_kw = Suppress(Keyword('EndIf')).setName('endif')
enum_kw = Suppress(Keyword('Enum')).setName('enum')
erase_kw = Suppress(Keyword('Erase')).setName('erase')
event_kw = Suppress(Keyword('Event')).setName('event')
exit_kw = Suppress(Keyword('Exit')).setName('exit')
for_kw = Suppress(Keyword('For')).setName('for')
friend_kw = Suppress(Keyword('Friend')).setName('friend')
function_kw = Suppress(Keyword('Function')).setName('function')
get_kw = Suppress(Keyword('Get')).setName('get')
global_kw = Suppress(Keyword('Global')).setName('global')
gosub_kw = Suppress(Keyword('GoSub')).setName('gosub')
goto_kw = Suppress(Keyword('GoTo')).setName('goto')
if_kw = Suppress(Keyword('If')).setName('if')
implements_kw = Suppress(Keyword('Implements')).setName('implements')
input_kw = Suppress(Keyword('Input')).setName('input')
let_kw = Suppress(Keyword('Let')).setName('let')
lock_kw = Suppress(Keyword('Lock')).setName('lock')
loop_kw = Suppress(Keyword('Loop')).setName('loop')
lset_kw = Suppress(Keyword('LSet')).setName('lset')
next_kw = Suppress(Keyword('Next')).setName('next')
on_kw = Suppress(Keyword('On')).setName('on')
open_kw = Suppress(Keyword('Open')).setName('open')
option_kw = Suppress(Keyword('Option')).setName('option')
print_kw = Suppress(Keyword('Print')).setName('print')
private_kw = Suppress(Keyword('Private')).setName('private')
public_kw = Suppress(Keyword('Public')).setName('public')
put_kw = Suppress(Keyword('Put')).setName('put')
raiseevent_kw = Suppress(Keyword('RaiseEvent')).setName('raiseevent')
redim_kw = Suppress(Keyword('ReDim')).setName('redim')
resume_kw = Suppress(Keyword('Resume')).setName('resume')
return_kw = Suppress(Keyword('Return')).setName('return')
rset_kw = Suppress(Keyword('RSet')).setName('rset')
seek_kw = Suppress(Keyword('Seek')).setName('seek')
select_kw = Suppress(Keyword('Select')).setName('select')
set_kw = Suppress(Keyword('Set')).setName('set')
static_kw = Suppress(Keyword('Static')).setName('static')
stop_kw = Suppress(Keyword('Stop')).setName('stop')
sub_kw = Suppress(Keyword('Sub')).setName('sub')
type_kw = Suppress(Keyword('Type')).setName('type')
unlock_kw = Suppress(Keyword('Unlock')).setName('unlock')
wend_kw = Suppress(Keyword('Wend')).setName('wend')
while_kw = Suppress(Keyword('While')).setName('while')
with_kw = Suppress(Keyword('With')).setName('with')
write_kw = Suppress(Keyword('Write')).setName('write')

any_kw = Suppress(Keyword("Any")).setName('any')
as_kw = Suppress(Keyword("As")).setName('as')
byref_kw = Suppress(Keyword("ByRef")).setName('byref')
byval_kw = Suppress(Keyword("ByVal")).setName('byval')
case_kw = Suppress(Keyword("Case")).setName('case')
each_kw = Suppress(Keyword("Each")).setName('each')
error_kw = Suppress(Keyword("Error")).setName('error')
else_kw = Suppress(Keyword("Else")).setName('else')
in_kw = Suppress(Keyword("In")).setName('in')
new_kw = Suppress(Keyword("New")).setName('new')
shared_kw = Suppress(Keyword("Shared")).setName('shared')
until_kw = Suppress(Keyword("Until")).setName('until')
withevents_kw = Suppress(Keyword("WithEvents")).setName('withevents')
write_kw = Suppress(Keyword("Write")).setName('write')
optional_kw = Suppress(Keyword("Optional")).setName('optional')
paramarray_kw = Suppress(Keyword("ParamArray")).setName('paramarray')
preserve_kw = Suppress(Keyword("Preserve")).setName('preserve')
spc_kw = Suppress(Keyword("Spc")).setName('spc')
step_kw = Suppress(Keyword('Step')).setName('step')
tab_kw = Suppress(Keyword("Tab")).setName('tab')
then_kw = Suppress(Keyword("Then")).setName('then')
to_kw = Suppress(Keyword("To")).setName('to')

statement_keyword = call_kw | case_kw | close_kw | const_kw | declare_kw \
    | defbool_kw | defbyte_kw | defcur_kw | defdate_kw | defdbl_kw \
    | defint_kw | deflng_kw | deflnglng_kw | deflngptr_kw | defobj_kw \
    | defsng_kw | defstr_kw | defvar_kw | dim_kw | do_kw | else_kw \
    | elseif_kw | end_kw | endif_kw | enum_kw | erase_kw | event_kw \
    | exit_kw | for_kw | friend_kw | function_kw | get_kw | global_kw \
    | gosub_kw | goto_kw | if_kw | implements_kw | input_kw | let_kw \
    | lock_kw | loop_kw | lset_kw | next_kw | on_kw | open_kw | option_kw \
    | print_kw | private_kw | public_kw | put_kw | raiseevent_kw | redim_kw \
    | resume_kw | return_kw | rset_kw | seek_kw | select_kw | set_kw \
    | static_kw | stop_kw | sub_kw | type_kw | unlock_kw | wend_kw | while_kw \
    | with_kw | write_kw

marker_keyword = any_kw | as_kw | byref_kw | byval_kw | case_kw | each_kw \
    | else_kw | in_kw | new_kw | shared_kw | until_kw | withevents_kw \
    | write_kw | optional_kw | paramarray_kw | preserve_kw | spc_kw | step_kw \
    | tab_kw | then_kw | to_kw

# OPERATOR

addressof_kw = Keyword('AddressOf').setName('addressof')
and_kw = Keyword('And').setName('and')
eqv_kw = Keyword('Eqv').setName('eqv')
imp_kw = Keyword('Imp').setName('imp')
is_kw = Keyword('Is').setName('is')
like_kw = Keyword('Like').setName('like')
mod_kw = Keyword('Mod').setName('mod')
not_kw = Keyword('Not').setName('not')
or_kw = Keyword('Or').setName('or')
typeof_kw = Keyword('TypeOf').setName('typeof')
xor_kw = Keyword('Xor').setName('xor')

operator_keyword = addressof_kw | and_kw | eqv_kw | imp_kw | is_kw | like_kw \
    | mod_kw | not_kw | or_kw | typeof_kw | xor_kw

# TYPES

# Numbers
integer = Word(nums).setName("integer") \
    .setParseAction(lambda r: Literal(Type.Integer, r[0]))

# Boolean
true_kw = Keyword('True')
false_kw = Keyword('False')

boolean = (true_kw | false_kw).setName("boolean") \
    .setParseAction(lambda r: Literal(Type.Boolean, r[0]))

# String
string = QuotedString(quoteChar='"', escQuote='""').setName("string") \
    .setParseAction(lambda r: Literal(Type.String, r[0]))


literal = (integer | boolean | string).setName("literal")
literal_kw = true_kw | false_kw

# RESERVED

reserved = statement_keyword | marker_keyword | operator_keyword | literal_kw

# IDENTIFIER

identifier_regex = r"(?:[a-zA-Z]|_[a-zA-Z])[a-zA-Z0-9_]*"
identifier = (~reserved + Regex(identifier_regex)).setName("identifier") \
    .setParseAction(lambda r: Identifier(r[0]))
identifier_keyword = Regex(identifier_regex).setName("identifier keyword") \
    .setParseAction(lambda r: Identifier(r[0]))

# Types
variable_type = Or(types) | identifier

#############
#  Grammar  #
#############

expression = Forward().setName("expression")
statement = Forward().setName("statement")


#################
#  Expressions  #
#################

# Member access
def __build_recursive_member_access(expr, pos, result):
    tokens = list(result)
    node = tokens[0]

    for token in tokens[1:]:
        if isinstance(token, Identifier):
            node = Get(node, token)
        elif isinstance(token, FunCall):
            assert(isinstance(token.function, Identifier))
            node = FunCall(Get(node, token.function), token.arguments)

    return node


orphan_function_call_paren = Forward().setName("orphan_function_call_paren")
member_access_token = dot + (orphan_function_call_paren | identifier_keyword)
leading_member_access_token = member_access_token \
    + FollowedBy(member_access_token)
member_access = (
    (orphan_function_call_paren + FollowedBy(member_access_token) | identifier)
    + pOptional((leading_member_access_token)[...] + dot + identifier_keyword)
).setParseAction(__build_recursive_member_access)


# Operator
def __build_binary_operator(expr, pos, result):
    tokens = result[0].asList()
    operator_symbols = tokens[1::2]
    operands = tokens[0::2]

    tree = operands[0]
    for operator_symbol, operand in zip(operator_symbols, operands[1:]):
        operator = BinaryOperator.build_operator(operator_symbol)
        tree = BinOp(operator, tree, operand)

    return tree


binary_operators = OPERATORS_MAP.get_precedence_list(__build_binary_operator)

# Function call
arguments_list_call = pOptional(delimitedList(expression)) \
    .setName("arguments_list_call") \
    .setParseAction(lambda r: ArgListCall(list(r)))
arguments_list_def = pOptional(delimitedList(identifier)) \
    .setName("arguments_list_def") \
    .setParseAction(lambda r: ArgListDef(list(r)))
orphan_function_call_paren << (identifier + lparen + arguments_list_call
                               + rparen) \
    .setName("orphan_function_call_paren") \
    .setParseAction(lambda r: FunCall(*r))
function_call_paren = (member_access + lparen + arguments_list_call + rparen) \
    .setName("function_call_paren") \
    .setParseAction(lambda r: FunCall(r[0], r[1]))
function_call_no_paren = (StringStart() + member_access + arguments_list_call
                          + StringEnd()) \
    .setName("function_call_no_paren") \
    .setParseAction(lambda r: FunCall(r[0], r[1]))

terminal = (literal | function_call_paren | member_access)
expression << infixNotation(
    terminal, binary_operators, lpar=lparen, rpar=rparen) \
    .setName('expression')
expression_statement = (function_call_no_paren | expression) \
    .setName("expression_statement")

##################
#  Declarations  #
##################

# Variable
variable_declaration = (dim_kw + identifier
                        + pOptional(as_kw + variable_type
                                    + pOptional(Suppress('=') + expression))) \
    .setParseAction(lambda r: VarDec(*r)) \
    .setName("var dec")
variable_declaration_with_constructor = \
    (dim_kw + identifier + as_kw + new_kw + variable_type) \
    .setParseAction(lambda r: VarDec(*r, new=True))

# TODO difference between Set and Let
variable_assignment = (pOptional(let_kw | set_kw) + identifier
                       + Suppress('=') + expression) \
    .setParseAction(lambda r: VarAssign(*r)) \
    .setName("var assign")

# Function
procedure_header = (sub_kw + identifier
                    + pOptional(lparen + arguments_list_def + rparen)) \
    .setParseAction(lambda r: ProcDefHeader(*r)) \
    .setName("proc header")
procedure_footer = (end_kw + sub_kw) \
    .setParseAction(lambda r: ProcDefFooter()) \
    .setName("proc footer")

function_header = (function_kw + identifier
                   + pOptional(lparen + arguments_list_def + rparen)) \
    .setParseAction(lambda r: FunDefHeader(*r)) \
    .setName("function header")
function_footer = (end_kw + function_kw) \
    .setParseAction(lambda r: FunDefFooter()) \
    .setName("function footer")


declarative_statement = (
    variable_declaration_with_constructor | variable_declaration
    | variable_assignment
    | procedure_header | procedure_footer
    | function_header | function_footer) \
    .setName("declaration")

############################
#  Loops and conditionals  #
############################

# For
for_header = (for_kw + identifier + Suppress('=') + expression
              + to_kw + expression + pOptional(step_kw + expression)) \
    .setParseAction(lambda r: ForHeader(*r))
for_footer = (next_kw + pOptional(identifier)) \
    .setParseAction(lambda r: ForFooter(*r))

loop_statement = for_header | for_footer

# If
if_header = (if_kw + expression + pOptional(then_kw)) \
    .setParseAction(lambda r: IfHeader(*r))
elseif_header = (elseif_kw + expression + pOptional(then_kw)) \
    .setParseAction(lambda r: ElseIfHeader(*r))
else_header = else_kw.setParseAction(lambda r: ElseHeader())
if_footer = (end_kw + if_kw).setParseAction(lambda r: IfFooter())

conditional_statement = if_header | elseif_header | else_header | if_footer


####################
#  Error handling  #
####################

on_error = (on_kw + error_kw + ((goto_kw + (integer | identifier))
                                | (resume_kw + next_kw))) \
    .setParseAction(lambda r: OnError(r[0]) if len(r) > 0 else OnError())
resume = resume_kw + pOptional(next_kw | (integer | identifier)) \
    .setParseAction(lambda r: Resume(r[0] if len(r) > 0 else None))
error = (error_kw + integer).setParseAction(lambda r: ErrorStatement(r[0]))

error_statement = on_error | resume | error

# Wrap up
statement <<= declarative_statement | loop_statement | conditional_statement \
    | expression_statement | error_statement

############
#  Parser  #
############

# Packrat parsing cache parsing results, improving speed
ParserElement.enablePackrat(cache_size_limit=128)


class Parser:
    """
    Syntactic parser used to transform a list of instructions into an abstract
    syntax tree.
    """
    __nested_blocks: List[PartialBlock]

    def __init__(self) -> None:
        pass

    def build_ast(self, instructions: List[Instruction], file_name: str = "") \
            -> AST:
        """
        Main method of the Parser. Parses each of the instructions using a
        pyparsing grammar, and then process the results to handle multiline
        structures such as loops, conditionals, etc.
        """
        self.__nested_blocks = [PartialBlock()]
        for instruction in instructions:
            if instruction.instruction.strip() != '':
                parsed_instruction = self.__parse_instruction(instruction)
                if isinstance(parsed_instruction, Statement):
                    self.__handle_statement(parsed_instruction)
                elif isinstance(parsed_instruction, BlockElement):
                    self.__handle_block_element(parsed_instruction)
                else:
                    assert (False)

        top_level_block = self.__nested_blocks.pop()
        main_sequence = Block(top_level_block.statements_blocks.pop())
        return main_sequence

    def __parse_instruction(self, instruction: Instruction)\
            -> Union[Statement, BlockElement]:
        """Use the pyparsing grammar to parse a single instruction."""
        try:
            parse_results = statement.parseString(
                instruction.instruction, parseAll=True)
        except ParseException as e:
            raise ParsingError(instruction.file_name, instruction.line_number,
                               str(e))

        return parse_results[0]

    def __handle_statement(self, statement: Statement) -> None:
        """
        Add a single-line statement to the last statements block of the
        innermost nest block.
        """
        current_block = self.__nested_blocks[-1]
        current_block.statements_blocks[-1].append(statement)

    def __handle_block_element(self, block_element: BlockElement) -> None:
        """
        Add a block element to the list of elements of the innermost partial
        block. Create it or add it to the statements of the parent block if
        needed.
        """
        if block_element.IsHeader:
            new_block = PartialBlock([block_element])
            self.__nested_blocks.append(new_block)
        else:
            self.__nested_blocks[-1].elements.append(block_element)
            if not block_element.IsFooter:
                self.__nested_blocks[-1].statements_blocks.append([])

        if block_element.IsFooter:
            complete_block = self.__nested_blocks.pop().build_block()
            parent_block = self.__nested_blocks[-1]
            parent_block.statements_blocks[-1].append(complete_block)

    @staticmethod
    def parse(content: str, file_name: Optional[str] = "") -> AST:
        """
        Parse the content of a file into an abstract syntax tree.

        :arg content: str content of the file
        :arg file_name: Optional name of the file
        :return: An AST representing the syntax of the file
        """
        parser = Parser()
        instructions = Preprocessor.preprocess(content, file_name)
        tree = parser.build_ast(instructions)

        return tree

    @staticmethod
    def parse_file(path: str) -> AST:
        """
        Parse a file into an abstract syntax tree.

        :arg path: Path of the file
        :return: An AST representing the syntax of the file
        """
        parser = Parser()
        instructions = Preprocessor.preprocess_file(path)
        tree = parser.build_ast(instructions)

        return tree
