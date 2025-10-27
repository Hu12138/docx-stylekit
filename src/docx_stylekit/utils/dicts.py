from copy import deepcopy


def deep_merge(base, override, *, replace_lists=True):
    """
    深度合并字典：
    - base 先复制
    - override 的值优先生效
    - 对 dict 做递归合并
    - 对 list 默认用 override 替换（replace_lists=True），否则拼接
    """
    if base is None:
        return deepcopy(override)
    if override is None:
        return deepcopy(base)

    if not isinstance(base, dict) or not isinstance(override, dict):
        return deepcopy(override)

    merged = deepcopy(base)
    for key, val in override.items():
        if key in merged and isinstance(merged[key], dict) and isinstance(val, dict):
            merged[key] = deep_merge(merged[key], val, replace_lists=replace_lists)
        elif (
            key in merged
            and isinstance(merged[key], list)
            and isinstance(val, list)
            and not replace_lists
        ):
            merged[key] = merged[key] + val
        else:
            merged[key] = deepcopy(val)
    return merged


__all__ = ["deep_merge"]
