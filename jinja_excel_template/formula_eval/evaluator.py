import ast
from typing import Any, Dict
from functools import lru_cache
from .functions import DEFAULT_FUNCTIONS


@lru_cache(maxsize=2048)
def parse_expr(expr: str):
    return ast.parse(expr, mode="eval")

class EvaluationError(Exception):
    pass


class Evaluator(ast.NodeVisitor):
    def __init__(self, context: Dict[str, Any] = None, functions: Dict[str, Any] = None):
        self.context = context or {}
        self.functions = {k.upper(): v for k, v in (functions or DEFAULT_FUNCTIONS).items()}

    def visit(self, node):
        if node is None:
            return None
        return super().visit(node)

    def visit_Expression(self, node: ast.Expression):
        return self.visit(node.body)

    def visit_BinOp(self, node: ast.BinOp):
        left = self.visit(node.left)
        right = self.visit(node.right)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
        if isinstance(node.op, ast.Mod):
            return left % right
        if isinstance(node.op, ast.Pow):
            return left ** right
        raise EvaluationError(f"Unsupported binary operator: {type(node.op)}")

    def visit_UnaryOp(self, node: ast.UnaryOp):
        operand = self.visit(node.operand)
        if isinstance(node.op, ast.UAdd):
            return +operand
        if isinstance(node.op, ast.USub):
            return -operand
        raise EvaluationError(f"Unsupported unary operator: {type(node.op)}")

    def visit_Num(self, node: ast.Num):
        return node.n

    def visit_Constant(self, node: ast.Constant):
        return node.value

    def visit_Name(self, node: ast.Name):
        if node.id in self.context:
            return self.context[node.id]
        # allow constants like True/False/None
        if node.id in ("True", "False", "None"):
            return eval(node.id)
        raise EvaluationError(f"Undefined name: {node.id}")

    def visit_Call(self, node: ast.Call):
        # only allow function names (no attribute calls)
        if isinstance(node.func, ast.Name):
            func_name = node.func.id.upper()
            if func_name not in self.functions:
                raise EvaluationError(f"Unsupported function: {func_name}")
            args = [self.visit(arg) for arg in node.args]
            # flatten list arguments if list passed
            func = self.functions[func_name]
            try:
                # try calling with context
                return func(args, self.context)
            except TypeError:
                return func(args)
        raise EvaluationError("Only named functions are allowed")

    def visit_List(self, node: ast.List):
        return [self.visit(el) for el in node.elts]

    def visit_Tuple(self, node: ast.Tuple):
        return [self.visit(el) for el in node.elts]

    def generic_visit(self, node):
        raise EvaluationError(f"Unsupported expression: {type(node)}")


def safe_eval(expr: str, context: Dict[str, Any] = None, functions: Dict[str, Any] = None):
    """
    Evaluate a restricted expression using AST. Only simple arithmetic operations and
    a small function set are allowed.
    """
    if expr is None:
        return None
    if isinstance(expr, (int, float, str)) and not isinstance(expr, str):
        return expr
    # allow formula strings starting with '=' commonly used in Excel (strip first char)
    if isinstance(expr, str) and expr.startswith("="):
        expr = expr[1:]
    # Pre-process Excel A1 range syntax, e.g. A1:B2 => RANGE("A1","B2")
    import re

    # Support Sheet!A1:A10 -> RANGE("Sheet","A1","A10")
    expr = re.sub(r"([A-Za-z0-9_]+)!([A-Za-z]+[0-9]+):([A-Za-z]+[0-9]+)", r'RANGE("\1","\2","\3")', expr)
    # Support A1:A10 -> RANGE("A1","A10")
    expr = re.sub(r"([A-Za-z]+[0-9]+):([A-Za-z]+[0-9]+)", r'RANGE("\1","\2")', expr)
    # Support Sheet!A1 -> GET("Sheet","A1")
    expr = re.sub(r"([A-Za-z0-9_]+)!([A-Za-z]+[0-9]+)", r'GET("\1","\2")', expr)
    try:
        parsed = parse_expr(expr)
    except SyntaxError as e:
        raise EvaluationError(f"Syntax error in expression: {expr}") from e
    evaluator = Evaluator(context=context, functions=functions)
    return evaluator.visit(parsed)
