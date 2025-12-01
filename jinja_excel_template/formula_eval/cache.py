from functools import lru_cache


@lru_cache(maxsize=1024)
def cached_eval_key(expr: str):
    return expr

