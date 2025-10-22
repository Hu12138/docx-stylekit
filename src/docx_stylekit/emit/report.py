import json

def print_diff_report(diffs, fmt="text"):
    if fmt == "json":
        print(json.dumps(diffs, ensure_ascii=False, indent=2))
        return
    for d in diffs:
        status = d["status"]
        path = d["path"]
        a = d["a"]
        b = d["b"]
        if status == "added":
            print(f"[+] {path}: {b}")
        elif status == "removed":
            print(f"[-] {path}: {a}")
        else:
            print(f"[±] {path}: {a}  -->  {b}")
