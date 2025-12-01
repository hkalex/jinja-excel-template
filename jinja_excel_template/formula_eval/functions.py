from dateutil.parser import parse as dt_parse

def average(args):
    if not args:
        return 0
    # If a single list argument is provided, use its items
    if len(args) == 1 and isinstance(args[0], (list, tuple)):
        src = args[0]
    else:
        src = args
    return sum(src) / len(src)

def concat(args):
    if len(args) == 1 and isinstance(args[0], (list, tuple)):
        src = args[0]
    else:
        src = args
    return "".join(str(a) for a in src)

def round_fn(args):
    # ROUND(x, ndigits)
    if len(args) == 0:
        return 0
    x = args[0]
    nd = int(args[1]) if len(args) > 1 else 0
    return round(x, nd)

def to_int(args):
    if not args:
        return 0
    return int(args[0])

def abs_fn(args):
    if not args:
        return 0
    return abs(args[0])

def max_fn(args):
    # flatten list if single list argument
    seq = args[0] if len(args) == 1 and isinstance(args[0], (list, tuple)) else args
    return max(seq)

def min_fn(args):
    seq = args[0] if len(args) == 1 and isinstance(args[0], (list, tuple)) else args
    return min(seq)

def date_fn(args):
    # DATE('YYYY-MM-DD') or DATE(year,month,day)
    if not args:
        return None
    if len(args) == 1 and isinstance(args[0], str):
        return dt_parse(args[0])
    if len(args) >= 3 and all(isinstance(a, (int, float)) for a in args[:3]):
        from datetime import datetime

        year = int(args[0])
        month = int(args[1])
        day = int(args[2])
        return datetime(year, month, day)
    raise ValueError("DATE requires either a string or (year,month,day)")

def range_fn(args, ctx=None):
    # args: ('A1', 'A3') or ('sheet','A1', 'A3')
    if ctx is None or '__range_resolver__' not in ctx:
        raise ValueError('Range resolver not provided in context')
    if len(args) == 2:
        start, end = args
        return ctx['__range_resolver__'](start, end)
    if len(args) == 3:
        sheet, start, end = args
        return ctx['__range_resolver__'](start, end, sheet)
    raise ValueError('RANGE expects 2 or 3 args')

def get_fn(args, ctx=None):
    if ctx is None or '__range_resolver__' not in ctx:
        raise ValueError('Range resolver not provided in context')
    if len(args) == 2:
        sheet, addr = args
        vals = ctx['__range_resolver__'](addr, addr, sheet)
        return vals[0] if vals else None
    if len(args) == 1:
        addr = args[0]
        vals = ctx['__range_resolver__'](addr, addr)
        return vals[0] if vals else None
    raise ValueError('GET expects 1 or 2 args')

DEFAULT_FUNCTIONS = {
    "SUM": lambda args: sum(args[0] if len(args)==1 and isinstance(args[0], (list,tuple)) else args),
    "AVERAGE": lambda args: average(args),
    "CONCAT": lambda args: concat(args),
    "ROUND": lambda args: round_fn(args),
    "INT": lambda args: to_int(args),
    "ABS": lambda args: abs_fn(args),
    "MAX": lambda args: max_fn(args),
    "MIN": lambda args: min_fn(args),
    "DATE": lambda args: date_fn(args),
    "RANGE": lambda args, ctx=None: range_fn(args, ctx=ctx),
    "GET": lambda args, ctx=None: get_fn(args, ctx=ctx),
}
