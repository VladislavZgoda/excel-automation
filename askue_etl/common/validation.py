def require_not_none[T](value: T | None, name: str) -> T:
    if value is None:
        raise ValueError(f"Требуется значение для {name}, получен None.")

    return value
