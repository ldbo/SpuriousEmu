"""Deobfuscation tools."""

from enum import Enum
from random import choices, shuffle
from statistics import mean
from string import digits, ascii_lowercase
from typing import Any, Dict, Set, Tuple
from typing import Type as tType

from .abstract_syntax_tree import *
from .compiler import Program
from .error import DeobfuscationError
from .interpreter import Resolver, Interpreter
from .reference import *
from .visitor import Visitor

# TODO add pure function evaluation
# TODO sharing variable names between functions causes issues because of how
#      the resolver works


class Deobfuscator(Visitor[AST]):
    """
    Class performing AST de-obfuscation.

    You can configure it using:
        - evaluation_level: see EvaluationLevel
        - rename_symbols: if True, rename symbols that appear to be obfuscated

    You can clean an AST with the deobfuscate method. To add support for a new
    AST node type you just need to implement the corresponding visit_ method,
    which should return an equivalent de-obfuscated node. If it returns None,
    standard deep de-obfuscated copy is performed.
    """

    class EvaluationLevel(Enum):
        """
        Used to specify what expressions should be evaluated : none of them,
        literal expressions or pure functions.
        """

        NONE = 0
        LITERAL = 1
        FUNCTION = 2

    REFERENCE_PREFIXES = {
        Environment: "Environment",
        Project: "Project",
        ClassModule: "Class",
        ProceduralModule: "Module",
        Variable: "var",
    }

    CALLABLE_PREFIXES = {FunDef: "Function", ProcDef: "Procedure"}

    __program: Program
    __resolver: Resolver
    __interpreter: Interpreter

    __new_names: Dict[tType[Reference], str]
    __references_count: Dict[tType[Reference], int]
    __callable_count: Dict[Union[tType[FunDef], ProcDef], int]

    __evaluation_level: "Deobfuscator.EvaluationLevel"
    rename_symbols: bool
    mangling_classifier: Optional["ManglingClassifier"]

    # Configuration

    def __init__(
        self,
        program: Program,
        evaluation_level: Union["Deobfuscator.EvaluationLevel", int] = 1,
        rename_symbols: bool = False,
        mangling_classifier: Optional["ManglingClassifier"] = None,
    ) -> None:
        self.__program = program
        self.__interpreter = Interpreter(program)
        self.__resolver = Resolver(self.__interpreter, program)

        self.__new_names = dict()
        self.__references_count = {
            reference_type: 0
            for reference_type in Deobfuscator.REFERENCE_PREFIXES.keys()
        }
        self.__callable_count = {
            callable_type: 0
            for callable_type in Deobfuscator.CALLABLE_PREFIXES.keys()
        }

        self.evaluation_level = evaluation_level
        self.rename_symbols = rename_symbols
        self.mangling_classifier = mangling_classifier

    @property
    def evaluation_level(self) -> "Deobfuscator.EvaluationLevel":
        return self.__evaluation_level

    @evaluation_level.setter
    def evaluation_level(
        self, level: Union["Deobfuscator.EvaluationLevel", int]
    ) -> None:
        if level == Deobfuscator.EvaluationLevel.FUNCTION:
            msg = "Pure functions deobfuscation is not supported yet"
            raise NotImplemented(msg)
        self.__evaluation_level = Deobfuscator.EvaluationLevel(level)

    # API

    def deobfuscate(self, ast: AST) -> AST:
        """
        Return the de-obfuscated version of an AST.
        """
        clean_ast = self.__deobfuscate(ast)
        assert isinstance(clean_ast, AST)

        return clean_ast

    def deobfuscate_module(self, name: str) -> AST:
        """
        Return the de-obfuscated version of a module.
        """
        try:
            ast = self.__program.asts[name]
        except KeyError:
            msg = f"Can't find module {name}"
            raise DeobfuscationError(msg)

        self.__resolver.jump(self.__resolver.resolve(Identifier(name)))
        clean_ast = self.__deobfuscate(ast)
        assert isinstance(clean_ast, AST)

        return clean_ast

    # General utility

    def __deobfuscate(self, elt: Any) -> Any:
        """
        Internal de-obfuscation method, which can handle AST, lists (whose
        elements are individually de-obfuscated) and any other type, which is
        returned as-is.
        """
        # list
        if isinstance(elt, list):
            return list(map(self.__deobfuscate, elt))

        # non AST
        if not isinstance(elt, AST):
            return elt

        try:
            # visitable AST
            return_value = self.visit(elt)

            if return_value is None:
                raise NotImplementedError("Back to default de-obfuscation.")

            return return_value
        except NotImplementedError:
            # non visitable AST: deep de-obfuscated copy
            return self.__deobfuscated_copy(elt)

    def __deobfuscated_copy(self, ast: AST, **new_attributes) -> AST:
        """
        Performs a deep copy of an AST node, de-obfuscating each of its public
        fields.

        :arg new_attributes: Use it to force the value of some fields
        :returns: An equivalent de-obfuscated AST
        """
        for attribute, value in ast.__dict__.items():
            if attribute.startswith("_") or attribute in new_attributes:
                continue

            new_attributes[attribute] = self.__deobfuscate(value)

        return type(ast)(**new_attributes)

    # Expressions evaluation

    def __is_evaluable(self, expression: Expression) -> bool:
        """
        Check if an expression can be evaluated by the deobfuscator, depending
        on the value of self.evaluation_level.
        """
        if self.evaluation_level == self.EvaluationLevel.NONE:
            return False
        elif self.evaluation_level == self.EvaluationLevel.LITERAL:
            return isinstance(expression, Literal)
        elif self.evaluation_level == self.EvaluationLevel.FUNCTION:
            return isinstance(expression, (Literal, FunCall))
        else:
            msg = f"Invalid evaluation level: {self.evaluation_level}"
            raise DeobfuscationError(msg)

    def __unroll_expression(
        self, expression: Expression, operator: Optional[str] = None
    ) -> List[Expression]:
        """
        Used to transform associative operator tree to list of expressions. In
        other words, remove parenthesis from an expression. For example, if used
        with (a + (b + c)) + (a + (b * c)), it would return [a, b, c, a, b*c].

        :arg expression: Expression to unroll
        :arg operator: Select the operator to unroll relatively to
        :returns: The recursively unrolled expression relatively to operator
        """
        if not isinstance(expression, BinOp):
            return [expression]
        elif operator is not None and operator != expression.operator:
            return [expression]
        else:
            unrolled_left = self.__unroll_expression(
                expression.left, expression.operator
            )
            unrolled_right = self.__unroll_expression(
                expression.right, expression.operator
            )

            return unrolled_left + unrolled_right

    def __evaluate_unrolled_expression(
        self, operator: str, exprs: List[Expression]
    ) -> Expression:
        """
        Evaluate consecutive evaluable elements of an unrolled expression with
        an associative operator. Return an equivalent expression.
        """
        # TODO work with tuples instead of list + copy
        # TODO add commutativity support
        assert len(exprs) >= 1
        if len(exprs) == 1:
            expr = exprs[0]
            return expr
        else:
            # Copy exprs
            exprs = list(exprs)
            right = exprs.pop()
            left = exprs[-1]

            left_evaluable = self.__is_evaluable(left)
            right_evaluable = self.__is_evaluable(right)

            if left_evaluable and right_evaluable:
                tmp_op = BinOp(operator, left, right)
                value = self.__interpreter.evaluate(tmp_op)
                exprs[-1] = Literal.from_value(value)
                return self.__evaluate_unrolled_expression(operator, exprs)
            else:
                new_left = self.__evaluate_unrolled_expression(operator, exprs)
                return BinOp(operator, new_left, right)

    @classmethod
    def __is_associative(cls, operator: str) -> bool:
        """Check if an operator is associative."""
        # TODO move to operator.py
        if operator in ("+", "*", "&", "And", "Or", "Eqv"):
            return True
        else:
            return False

    # Symbol demangling
    def __is_mangled(self, name: str) -> bool:
        """
        If no classifier is specified, return True. Else, consider name as an
        underscore-separated list of words. It is mangled if one of this is
        mangled, ie. it has non-ending digits or its non-digit prefix is
        classified as mangled by the classifier.
        """
        if self.mangling_classifier is None:
            return True

        for word in name.lower().split("_"):
            for i in range(len(word) - 1):
                if word[i] in digits and word[i + 1] in ascii_lowercase:
                    return True
                if word[i + 1] in digits and word[i] in ascii_lowercase:
                    first_digit_position = i + 1

            if word[: i + 1] in self.mangling_classifier:
                return True

        return False

    def __craft_name(self, reference: Reference) -> str:
        """
        Craft a new name for the given reference, using the format
        {ReferenceType}_XX. See Deobfuscator.REFERENCE_PREFIXES.
        """
        count = self.__references_count[type(reference)] + 1
        self.__references_count[type(reference)] = count
        prefix = self.REFERENCE_PREFIXES[type(reference)]
        new_name = f"{prefix}_{count}"

        return new_name

    def __craft_arg(self, argument_count) -> str:
        """
        Craft a new argument name, using the format arg_XX.
        """
        return f"arg_{argument_count}"

    def __craft_callable_name(
        self, reference: FunctionReference, callable: Union[FunDef, ProcDef]
    ) -> str:
        """Craft a new function or procedure name, using the format
        [Function|Procedure]_XX. See Deobfuscator.CALLABLE_PREFIXES.
        """
        t = type(callable)
        count = self.__callable_count[t] + 1
        self.__callable_count[t] = count

        return f"{self.CALLABLE_PREFIXES[t]}_{count}"

    def __rename_identifier(self, identifier: Identifier) -> Identifier:
        """
        Craft a new name for the identifier using __craft_name
        """
        try:
            resolution = self.__resolver.resolve(identifier)
        except ResolutionError:
            return Identifier(identifier.name)

        try:
            new_name = self.__new_names[resolution]
        except KeyError:
            if self.__is_mangled(identifier.name):
                new_name = self.__craft_name(resolution)
            else:
                new_name = identifier.name

            self.__new_names[resolution] = new_name

        return self.__deobfuscated_copy(identifier, name=new_name)

    def __rename_get(self, get: Get) -> Identifier:
        """Only rename the top parent node."""
        # TODO Rename recursively
        new_child = self.__deobfuscated_copy(get.child, name=get.child.name)

        if isinstance(get.parent, (FunCall, Get)):
            return self.__deobfuscated_copy(get, child=new_child)
        else:
            new_parent = self.visit(get.parent)
            return self.__deobfuscated_copy(
                get, parent=new_parent, child=new_child
            )

    def __rename_callable_definition(
        self, definition: Union[ProcDef, FunDef]
    ) -> Union[ProcDef, FunDef]:
        """
        Handle first the function name and arguments, to correctly rename them,
        and then make resolver jump to function.
        """
        # TODO Use generic type to ensure type(definition) is type(return)
        function_reference = self.__resolver.resolve(definition.name)

        if self.__is_mangled(definition.name.name):
            new_name = self.__craft_callable_name(
                function_reference, definition
            )
        else:
            new_name = definition.name.name

        self.__new_names[function_reference] = new_name
        arguments = self.visit(definition.arguments)

        # Update return variable
        if isinstance(definition, FunDef):
            return_reference = function_reference.get_child(
                function_reference.name
            )
            self.__new_names[return_reference] = new_name

        self.__references_count[Variable] = 0
        self.__resolver.jump(function_reference)
        clean_definition = self.__deobfuscated_copy(
            definition, name=Identifier(new_name), arguments=arguments
        )
        self.__resolver.jump_back()

        return clean_definition

    def __rename_arg_list_def(self, arg_list: ArgListDef) -> ArgListDef:
        """
        Argument and local variable references do not use distinct types, so
        this enforces the use of __craft_arg instead of __craft_name.
        """
        arguments_count = 0
        args = []
        for arg_identifier in arg_list.args:
            arguments_count += 1

            if self.__is_mangled(arg_identifier.name):
                new_name = self.__craft_arg(arguments_count)
            else:
                new_parent = arg_identifier.name

            reference = self.__resolver.resolve(arg_identifier)
            self.__new_names[reference] = new_name
            args.append(Identifier(new_name))

        return self.__deobfuscated_copy(arg_list, args=args)

    # visit_ methods

    def visit_BinOp(self, bin_op: BinOp) -> AST:
        clean_bin_op = self.__deobfuscated_copy(bin_op)

        operator = clean_bin_op.operator
        if not self.__is_associative(operator):
            return clean_bin_op

        unrolled_bin_op = self.__unroll_expression(clean_bin_op)

        return self.__evaluate_unrolled_expression(operator, unrolled_bin_op)

    def visit_Identifier(self, identifier: Identifier) -> AST:
        if self.rename_symbols:
            return self.__rename_identifier(identifier)

    def visit_Get(self, get: Get) -> AST:
        if self.rename_symbols:
            return self.__rename_get(get)

    def visit_ArgListDef(self, arg_list: ArgListDef) -> AST:
        if self.rename_symbols:
            return self.__rename_arg_list_def(arg_list)

    def visit_FunDef(self, fun_def: FunDef) -> AST:
        if self.rename_symbols:
            return self.__rename_callable_definition(fun_def)

    def visit_ProcDef(self, proc_def: ProcDef) -> AST:
        if self.rename_symbols:
            return self.__rename_callable_definition(proc_def)


class ManglingClassifier:
    """
    Allows to detect mangled names based on Markov chain model of the supposedly
    used language.

    First use the constructor, or the train method, to train the model. Then,
    the `in` operator or the is_mangled method can be used to determine if a
    word is mangled.

    This classifier is only made to work with lower case letter-only words, ie.
    words matching the [a-z]* regex.
    """

    FrequencyMatrixType = Dict[Tuple[str, str], float]

    __n: int
    __false_negative_rate: float
    __frequency_matrix: "ManglingClassifier.FrequencyMatrixType"
    __ngrams: Set[str]
    __threshold: float

    def __init__(
        self,
        n: int,
        training_data: Union[
            List[str], "ManglingClassifier.FrequencyMatrixType"
        ],
        false_negative_rate: float,
    ) -> None:
        """
        Train a Markov classifier. First, compute the frequency matrix of the
        hidden model : it holds the probability of having a character given its
        n previous neighbours. Then, find the threshold parameter ensuring the
        false negative rate on a test set extracted from training_data.

        :arg n: Number of previous characters to take into account
        :arg training_data: List of words (must be big enough, tens of thousands
         for example) to train
        :arg false_negative_rate: Maximum false negative rate allowed on a test
        set extracted from the training data
        """
        self.n = n
        self.false_negative_rate = false_negative_rate

        self.train(training_data)

    def train(self, dictionnary: List[str]) -> None:
        """
        Build the frequency matrix of the hidden model and set the threshold.
        """
        test_set, training_set = self.__extract_training_test(dictionnary, 0.1)

        self.__build_frequency_matrix(training_set)
        self.__set_threshold(test_set)

    def word_probability(self, word: str) -> float:
        """
        Compute the mean probability of the ngram pairs of the word.
        """
        reduced_length = len(word) - self.n

        if reduced_length < 0:
            return 1.0
        elif reduced_length == 0:
            return 1.0 if word in self.__ngrams else 0.0
        else:
            probability_sum = sum(
                map(
                    lambda pair: self.__frequency_coefficient(*pair),
                    self.__ngram_pairs(word.lower()),
                )
            )
            return probability_sum / reduced_length

    def is_mangled(self, word: str) -> bool:
        """
        Check if a word appears to be mangled by comparing its probability with
        the threshold compute so that the training language has a false negative
        rate lower than false_negative_rate.


        :returns: True if the word seems to be mangled
        """
        return self.word_probability(word) < self.__threshold

    def __contains__(self, word: str) -> bool:
        """
        Check if a word is mangled using is_mangled.
        """
        return self.is_mangled(word)

    def __extract_training_test(
        self, sample: List[Any], test_rate: float
    ) -> Tuple[List[Any], List[Any]]:
        """
        Split a sample into a training and a test dataset.
        """
        sample_copy = list(sample)
        shuffle(sample_copy)
        training_length = int(test_rate * len(sample))

        return sample_copy[:training_length], sample_copy[training_length:]

    def __ngram_pairs(self, word):
        """Generator producing the list of (w[i]...w[i+n-1], w[i+n]) pairs."""
        for i in range(len(word) - self.n):
            end = i + self.n
            yield word[i:end], word[end]

    def __build_frequency_matrix(self, training_set: List[str]) -> None:
        """
        Extract statistical information from a list of strings: the set of
        ngrams and the frequency matrix.
        """
        ngram_pairs_counts: Dict[Tuple[str, str], int] = dict()
        ngram_counts: Dict[str, int] = dict()

        for word in training_set:
            for ngram, next_char in self.__ngram_pairs(word):
                ngram_counts[ngram] = ngram_counts.get(ngram, 0) + 1
                ngram_pairs_counts[ngram, next_char] = (
                    ngram_pairs_counts.get((ngram, next_char), 0) + 1
                )

        self.__frequency_matrix = {
            ngram_pair: pair_count / ngram_counts[ngram_pair[0]]
            for ngram_pair, pair_count in ngram_pairs_counts.items()
        }
        self.__ngrams = set(ngram_counts.keys())

    def __mean_sample_probability(self, sample: List[str]) -> float:
        """
        Return the mean probability of a list of words, using word_probability.
        """
        return mean(map(self.word_probability, sample))

    def __negative_rate(self, sample: List[str]) -> float:
        """
        Return the negative rate of is_mangled applied on a list of words.
        """
        return len(list(filter(lambda w: w in self, sample))) / len(sample)

    def __set_threshold(self, test_set: List[str]) -> None:
        """
        Find a threshold such that
        P(w in test_set: w not in self) < false_negative_rate.
        """
        self.__threshold = self.__mean_sample_probability(test_set)

        while self.__negative_rate(test_set) >= self.false_negative_rate:
            self.__threshold /= 2

    def __frequency_coefficient(self, ngram: str, next_char: str) -> float:
        """
        Return the probability of having next_char after ngram.
        """
        return self.__frequency_matrix.get((ngram, next_char), 0.0)
