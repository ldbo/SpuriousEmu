"""Pseudo-compilation stage: allow to extract references and function body."""

from dataclasses import dataclass
from importlib.util import spec_from_file_location, module_from_spec
from inspect import getmembers, isclass, isfunction
from itertools import chain
from pathlib import Path
from types import ModuleType
from typing import Dict, Any, List

from . import reference
from .abstract_syntax_tree import *
from .error import ResolutionError
from .function import ExternalFunction, InternalFunction
from .side_effect import Memory
from .syntax import Parser
from .vba_class import Class
from .visitor import Visitor


@dataclass
class Program:
    """
    Represent a statically analysed program : a set of references, with a
    memory containing already initialized values.
    """
    memory: Memory
    environment: reference.Environment

    def to_dict(self) -> Dict[str, Any]:
        d = dict()
        d['memory'] = {'functions': list(self.memory._functions.keys()),
                       'classes': list(self.memory._classes.keys())}
        d['environment'] = self.environment.to_dict()
        return d


class Compiler(Visitor):
    """
    Class used for references extraction. You can analyse several modules in a
    row by using add_module multiple times, and can then retrieve the built
    references and memory by accessing the corresponding properties. To start a
    new analysis, use reset.
    """
    __environment: reference.Environment
    __current_reference: reference.Reference
    __memory: Memory

    @staticmethod
    def module_type_from_name(file_name: str):
        """
        Determine the module type from the extension of the file, default to a
        procedural module.
        """
        extension = file_name.split('.')[-1]
        if extension == 'cls':
            return reference.ClassModule
        elif extension == 'bas' or extension == 'vbs':
            return reference.ProceduralModule
        else:
            raise RuntimeError('File {file_name} has unkown extension')

    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        """Reset the state of the compiler, erasing references and memory."""
        self.__environment = reference.Environment()
        self.__current_reference = self.__environment
        self.__memory = Memory()

    def add_project(self, project_name: str) -> None:
        """
        Add a project to the current program and use it as the parent project
        for the following operations.
        """
        project = self.__environment.build_child(reference.Project,
                                                 name=project_name)
        self.__current_reference = project

    def add_module(self, ast: AST, module_type,
                   module_name: str = None) -> None:
        """
        Add a module to the current project, analysing the AST of the module.
        """
        if isinstance(self.__current_reference, reference.Module):
            self.__current_reference = self.__current_reference.parent
        elif type(self.__current_reference) is reference.Project:
            pass
        elif type(self.__current_reference) is reference.Environment:
            try:
                project = self.__environment.get_child("Default")
            except ResolutionError:
                project = self.__environment.build_child(
                    reference.Project,
                    name="Default")

            self.__current_reference = project

        module = self.__current_reference.build_child(
            module_type, name=module_name)

        if module_type is reference.ClassModule:
            self.__memory.classes[str(module)] = Class([])

        self.__current_reference = module
        self.visit(ast)

    def load_host_project(self, project_path: str) -> None:
        """Load a Python package as a VBA host project."""
        path = Path(project_path)
        assert(path.is_dir())
        project_name = path.name

        project = reference.Project(project_name)
        self.__environment.add_child(project)
        self.__current_reference = project

        for module_path in path.glob("*.py"):
            module_name = f"{project_name}.{module_path.name[:-3]}"
            module_spec = spec_from_file_location(module_name,
                                                  module_path.absolute())
            module = module_from_spec(module_spec)
            module_spec.loader.exec_module(module)

            self.load_host_module(module)

        self.__current_reference = self.__current_reference.parent

    def load_host_module(self, module: ModuleType) -> None:
        """
        Extract objects defined in an already loaded Python module and add them
        to the current program.
        """
        def locally_defined(predicate):
            def decorated_predicate(member):
                return predicate(member) \
                    and member.__module__ == module.__name__

            return decorated_predicate

        module_name = module.__name__.split('.')[-1]

        classes = getmembers(module, locally_defined(isclass))
        assert(len(classes) in (0, 1))
        if len(classes) == 0:
            self.__current_reference = self.__current_reference.build_child(
                reference.ProceduralModule, module_name)
            functions_holder = module
        else:
            self.__current_reference = self.__current_reference.build_child(
                reference.ClassModule, module_name)
            class_name = classes[0][0]
            py_class = getattr(module, class_name)
            functions_holder = py_class

            for variable in py_class.variables:
                self.__current_reference.build_child(
                    reference.Variable,
                    variable,
                    extent=reference.Variable.Extent.Object
                )

            vba_class = Class(py_class.variables)
            self.__memory.classes[str(self.__current_reference)] = vba_class

        for name, function in getmembers(
                functions_holder, locally_defined(isfunction)):
            self.load_host_function(function)

        self.__current_reference = self.__current_reference.parent

    def load_host_function(self, function: ExternalFunction.Signature) \
            -> None:
        typed_function = ExternalFunction.from_function(function)
        name = typed_function.name

        function = self.__current_reference.build_child(
            reference.FunctionReference, name=name)
        self.__memory.add_function(str(function), typed_function)

    @property
    def program(self):
        """Return the compiled program."""
        return Program(self.__memory, self.__environment)

    def __parse_callable(self, definition: Union[FunDef, ProcDef]) -> None:
        # Extract information
        name = definition.name.name
        if definition.arguments is None:
            args = []
        else:
            args = list(map(lambda t: t.name, definition.arguments.args))
        body = definition.body

        # Build reference and memory representation
        function_ref = reference.FunctionReference(name)
        self.__current_reference.add_child(function_ref)
        function_object = InternalFunction(name, args, body)

        # Add them
        self.__memory.add_function(str(function_ref), function_object)

        self.__current_reference = function_ref

        for arg in args:
            self.__try_add_variable(arg)

        if type(definition) is FunDef:
            self.__try_add_variable(name)

        self.__try_add_variable(name)
        self.visit_Block(definition)
        self.__current_reference = self.__current_reference.parent

    def __try_add_variable(self, name: str) -> None:
        assert(isinstance(self.__current_reference,
                          (reference.Module, reference.FunctionReference)))

        try:
            self.__current_reference.get_child(name)
            return
        except ResolutionError:
            pass

        if isinstance(self.__current_reference, reference.ProceduralModule):
            extent = reference.Variable.Extent.Module
        elif isinstance(self.__current_reference, reference.ClassModule):
            extent = reference.Variable.Extent.Object
            vba_class = self.__memory.classes[str(self.__current_reference)]
            vba_class.variables.append(name)
        elif isinstance(self.__current_reference, reference.FunctionReference):
            extent = reference.Variable.Extent.Procedure

        self.__current_reference.build_child(
            reference.Variable,
            name=name,
            extent=extent)

    def visit_Block(self, block: Block) -> None:
        for statement in block.body:
            self.visit(statement)

    def visit_VarDec(self, var_dec: VarDec) -> None:
        self.__try_add_variable(var_dec.identifier.name)

    def visit_VarAssign(self, var_assign: VarAssign) -> None:
        if type(var_assign.variable) is Identifier:
            name = var_assign.variable.name
            self.__try_add_variable(name)

    def visit_FunDef(self, fun_def: FunDef) -> None:
        self.__parse_callable(fun_def)

    def visit_ProcDef(self, proc_def: ProcDef) -> None:
        self.__parse_callable(proc_def)

    def visit_FunCall(self, fun_call: FunCall) -> None:
        pass

    def visit_If(self, if_block: If) -> None:
        self.visit_Block(if_block)
        for elseif in if_block.elsifs:
            self.visit_Block(elseif)
        if if_block.else_block is not None:
            self.visit_Block(if_block.else_block)

    def visit_For(self, for_loop: For) -> None:
        self.__try_add_variable(for_loop.counter.name)
        self.visit_Block(for_loop)

    def visit_OnError(self, on_error: OnError) -> None:
        pass

    def visit_Resume(self, resume: Resume) -> None:
        pass

    def visit_ErrorStatement(self, error: ErrorStatement) -> None:
        pass

    @staticmethod
    def compile_file(path: str) -> Program:
        """Parse a file and then compile it, using Office extension."""
        return Compiler.compile_files([path])

    @staticmethod
    def compile_files(paths: List[str]) -> Program:
        """
        Parse and compile a list of files with Office extensions (cls and bas).
        """
        compiler = Compiler()

        for path in paths:
            ast = Parser.parse_file(path)
            module_type = Compiler.module_type_from_name(path)
            compiler.add_module(ast, module_type,
                                Path(path).name.split('.')[0])

        return compiler.program

    @staticmethod
    def compile_script(path: str) -> Program:
        """Compile a single file script, with vbs extension."""
        compiler = Compiler()
        ast = Parser.parse_file(path)
        module_type = reference.ProceduralModule
        compiler.add_module(ast, module_type, Path(path).name.split('.')[0])

        return compiler.program

    @staticmethod
    def compile_project(project_path: str) -> Program:
        """
        Recursively compile a directory, only taking into accound cls and bas
        files.
        """
        path = Path(project_path)
        assert(path.is_dir())
        paths = list(chain(list(path.rglob(f'*.{ext}'))
                           for ext in ('cls', 'bas')))
        paths = [str(file.absolute()) for glob in
                 (path.rglob(f"*.{ext}") for ext in ('cls', 'bas'))
                 for file in glob]

        return Compiler.compile_files(paths)
