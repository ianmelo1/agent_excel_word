import time
from functools import wraps


def rate_limit(max_per_minute):
    """Decorator para limitar chamadas à API"""
    min_interval = 60.0 / max_per_minute
    last_called = [0.0]

    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            elapsed = time.time() - last_called[0]
            left_to_wait = min_interval - elapsed

            if left_to_wait > 0:
                print(f"⏳ Aguardando {left_to_wait:.1f}s (rate limit)...")
                time.sleep(left_to_wait)

            ret = func(*args, **kwargs)
            last_called[0] = time.time()
            return ret

        return wrapper

    return decorator