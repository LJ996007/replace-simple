"""界面、规则文件和替换引擎共用的规则语义。"""

DELETE_MARKER = "〈删除原文〉"


def normalize_rule_rows(rows):
    rules = []
    seen = set()
    for row in rows:
        if not isinstance(row, (list, tuple)):
            continue
        old = "" if not row or row[0] is None else str(row[0]).strip()
        new = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
        if old and (old, new) not in seen:
            seen.add((old, new))
            rules.append((old, new))
    return rules


def rule_conflicts(rows):
    """返回同一原文有多个有效替换值的原始行号（从 1 开始）。"""
    by_old = {}
    for number, row in enumerate(rows, 1):
        for old, new in normalize_rule_rows([row]):
            if new:
                by_old.setdefault(old, {}).setdefault(new, []).append(number)
    return {old: sorted(n for numbers in values.values() for n in numbers)
            for old, values in by_old.items() if len(values) > 1}


def validate_rules(rows):
    conflicts = rule_conflicts(rows)
    if conflicts:
        raise ValueError("同一原文存在不同替换值，请修改冲突行：\n" + "\n".join(
            f"{old}：第 {'、'.join(map(str, numbers))} 行" for old, numbers in conflicts.items()
        ))
