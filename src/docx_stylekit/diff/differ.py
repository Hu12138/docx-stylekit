def dict_diff(a: dict, b: dict, path=""):
    """
    结构化对比：返回差异列表 [{"path":"x.y.z","a":valA,"b":valB,"status":"added|removed|changed"}]
    """
    diffs = []
    a_keys = set(a.keys())
    b_keys = set(b.keys())
    for k in sorted(a_keys - b_keys):
        diffs.append({"path": f"{path}.{k}" if path else k, "a": a[k], "b": None, "status": "removed"})
    for k in sorted(b_keys - a_keys):
        diffs.append({"path": f"{path}.{k}" if path else k, "a": None, "b": b[k], "status": "added"})
    for k in sorted(a_keys & b_keys):
        va, vb = a[k], b[k]
        cur_path = f"{path}.{k}" if path else k
        if isinstance(va, dict) and isinstance(vb, dict):
            diffs.extend(dict_diff(va, vb, cur_path))
        elif va != vb:
            diffs.append({"path": cur_path, "a": va, "b": vb, "status": "changed"})
    return diffs
