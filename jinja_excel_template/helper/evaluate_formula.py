from typing import Dict, Any
from jinja_excel_template.formula_eval import safe_eval, DEFAULT_FUNCTIONS
from jinja_excel_template.formula_eval.cache import cached_eval_key


def evaluate_formula(expr: str, context: Dict[str, Any] = None):
    """Evaluate a formula expression safely using the evaluator and caching."""
    if expr is None:
        return None
    cache_key = cached_eval_key(expr)
    # Right now, cached_eval_key is identity; we parse each time to keep implementation simple
    try:
        return safe_eval(expr, context=context, functions=DEFAULT_FUNCTIONS)
    except Exception as e:
        raise
