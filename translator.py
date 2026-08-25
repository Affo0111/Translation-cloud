#!/usr/bin/env python3
"""
表格翻译系统 - 基于自定义语法的订单定制项翻译工具

功能：根据学习模板中的规则，自动翻译和修改订单中的"定制项"字段，
     并根据条件修改SKU。

规则语法支持5种操作符：^（SKU变化）、!（删除）、=（转换）、++（指定位置添加）、+（行末添加）
内置修饰符：[]（范围限定）、:（键值分隔）、&（逻辑与）、|（逻辑或）、;（规则分隔）

用法：
    python translator.py --template template.xlsx --orders orders.xlsx --output result.xlsx
"""

import argparse
import logging
import os
import re
import sys
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Union

import pandas as pd
import openpyxl

# ──────────────────────────────────────────────
# 日志配置
# ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("translator")

# ── 版本号（用于 Streamlit Cloud 确认部署版本）──
__version__ = "2.5.0-styled"


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第一部分：转义字符处理                               ║
# ╚════════════════════════════════════════════════════════════════════════════╝

# 转义序列 → 占位符的映射表
# 占位符使用Unicode私有区字符，避免与正常内容冲突
_ESCAPE_MAP = {
    "\\\\": "\ue000",  # \\ → 反斜杠占位符
    "\\[":  "\ue001",  # \[ → 字面量 [
    "\\]":  "\ue002",  # \] → 字面量 ]
    "\\:":  "\ue003",  # \: → 字面量 :
    "\\&":  "\ue004",  # \& → 字面量 &
    "\\|":  "\ue005",  # \| → 字面量 |
    "\\;":  "\ue006",  # \; → 字面量 ;
    "\\,":  "\ue007",  # \, → 字面量 ,
}

# 反向映射
_UNESCAPE_MAP = {v: k[1] for k, v in _ESCAPE_MAP.items()}


def _escape_special(s: str) -> str:
    """将转义序列替换为占位符，使后续解析不受特殊字符干扰。"""
    result = s
    # 按长度降序排列，先替换多字符的转义序列
    for seq, placeholder in sorted(_ESCAPE_MAP.items(), key=lambda x: -len(x[0])):
        result = result.replace(seq, placeholder)
    return result


def _unescape(s: str) -> str:
    """将占位符恢复为原始字面量字符。"""
    result = s
    for placeholder, char in _UNESCAPE_MAP.items():
        result = result.replace(placeholder, char)
    return result


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第二部分：条件表达式                                 ║
# ╚════════════════════════════════════════════════════════════════════════════╝

@dataclass
class Condition:
    """条件表达式基类。"""

    def matches_line(self, line: str) -> bool:
        """检查单行定制项是否满足此条件。"""
        raise NotImplementedError

    def exists_in_lines(self, lines: List[str]) -> bool:
        """检查定制项行列表中是否存在满足此条件的行（用于 ^ 操作符）。"""
        raise NotImplementedError


@dataclass
class KeyValueCondition(Condition):
    """
    基本键值条件：[key:value] 或 [key]
    - 有 : 时，精确匹配键和值
    - 无 : 时，只匹配键（任意值）
    """
    key: str
    value: Optional[str]  # None 表示只匹配键

    def matches_line(self, line: str) -> bool:
        line = line.strip()
        if self.value is None:
            # 匹配键名或值：行以 "key:" 开头 或 值等于 self.key
            if ":" in line:
                line_key, line_value = line.split(":", 1)
                return (line_key.strip() == self.key or
                        line_value.strip() == self.key)
            return line == self.key  # 没有冒号，整行等于键名
        else:
            # 精确匹配键值对（分别 strip 键和值以容忍多余空格）
            if ":" in line:
                line_key, line_value = line.split(":", 1)
                return line_key.strip() == self.key and line_value.strip() == self.value
            return False

    def match_type(self, line: str) -> Optional[str]:
        """返回匹配类型：'key'（键匹配）、'value'（值匹配）、'both'（键值均匹配）、None（不匹配）。"""
        line = line.strip()
        if self.value is not None:
            if line == f"{self.key}:{self.value}":
                return "both"
            return None
        if ":" in line:
            line_key, line_value = line.split(":", 1)
            k_match = line_key.strip() == self.key
            v_match = line_value.strip() == self.key
            if k_match and v_match:
                return "key"  # 键优先
            elif k_match:
                return "key"
            elif v_match:
                return "value"
            return None
        if line == self.key:
            return "key"
        return None

    def exists_in_lines(self, lines: List[str]) -> bool:
        return any(self.matches_line(ln) for ln in lines)


@dataclass
class AndCondition(Condition):
    """逻辑与：所有子条件必须同时满足。"""
    children: List[Condition]

    def matches_line(self, line: str) -> bool:
        return all(c.matches_line(line) for c in self.children)

    def exists_in_lines(self, lines: List[str]) -> bool:
        # 所有子条件都必须在定制项中存在（可能在不同行）
        for child in self.children:
            if not child.exists_in_lines(lines):
                return False
        return True


@dataclass
class OrCondition(Condition):
    """逻辑或：满足任一子条件即可。"""
    children: List[Condition]

    def matches_line(self, line: str) -> bool:
        return any(c.matches_line(line) for c in self.children)

    def exists_in_lines(self, lines: List[str]) -> bool:
        return any(c.exists_in_lines(lines) for c in self.children)


def _parse_bracket_content(content: str) -> KeyValueCondition:
    """
    解析 [...] 内部的内容，构造 KeyValueCondition。
    例如：
        "尺寸:L" → KeyValueCondition(key="尺寸", value="L")
        "Name"   → KeyValueCondition(key="Name", value=None)
        "无:XYZ-NCT Larkspur" → KeyValueCondition(key="无", value="XYZ-NCT Larkspur")
    """
    content = content.strip()
    if ":" in content:
        # 只按第一个冒号分割（值中可能包含冒号，但在我们的场景中值通常不含冒号）
        key, value = content.split(":", 1)
        return KeyValueCondition(key=key.strip(), value=value.strip())
    else:
        return KeyValueCondition(key=content.strip(), value=None)


def parse_condition(expr: str) -> Condition:
    """
    解析条件表达式字符串，构造条件树。

    优先级：[] > & > |
    即先用 | 分割或组，再用 & 分割与条件。

    例如：
        "[A]&[B]|[C]" → OrCondition(AndCondition(A,B), C)
        "[A]|[B]&[C]" → OrCondition(A, AndCondition(B,C))
    """
    expr = expr.strip()

    # 1. 按 | 分割（或组），但只分割不在 [...] 内部的 |
    or_parts = _split_by_operator(expr, "|")
    if len(or_parts) > 1:
        return OrCondition([parse_condition(part) for part in or_parts])

    # 2. 按 & 分割（与组），但只分割不在 [...] 内部的 &
    and_parts = _split_by_operator(expr, "&")
    if len(and_parts) > 1:
        return AndCondition([parse_condition(part) for part in and_parts])

    # 3. 基本条件：[...]
    bracket_content = _extract_brackets(expr)
    if bracket_content is None:
        raise ValueError(f"无法解析条件表达式: {expr!r}")
    return _parse_bracket_content(bracket_content)


def _split_by_operator(expr: str, op: str) -> List[str]:
    """
    按操作符分割表达式，但忽略 [...] 内部的操作符。
    """
    parts = []
    depth = 0
    current = []
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == "[":
            depth += 1
            current.append(ch)
        elif ch == "]":
            depth -= 1
            current.append(ch)
        elif depth == 0 and expr[i:i + len(op)] == op:
            parts.append("".join(current).strip())
            current = []
            i += len(op)
            continue
        else:
            current.append(ch)
        i += 1
    parts.append("".join(current).strip())
    return [p for p in parts if p]  # 过滤空字符串


def _extract_brackets(expr: str) -> Optional[str]:
    """
    从表达式中提取最外层 [...] 的内容。
    例如 "[尺寸:L]" → "尺寸:L"
         "[Name]"   → "Name"
    """
    expr = expr.strip()
    if expr.startswith("[") and expr.endswith("]"):
        inner = expr[1:-1]
        # 验证括号配对
        if _brackets_balanced(inner):
            return inner
    return None


def _brackets_balanced(s: str) -> bool:
    """检查字符串中 [] 是否配对。"""
    depth = 0
    for ch in s:
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第三部分：规则解析                                   ║
# ╚════════════════════════════════════════════════════════════════════════════╝

class RuleParseError(Exception):
    """规则解析错误。"""
    pass


@dataclass
class Rule:
    """规则基类。"""
    operator: str  # "^", "!", "=", "++", "+"
    original: str  # 原始规则字符串

    def apply(self, sku: str, lines: List[str]) -> Tuple[str, List[str]]:
        """执行规则，返回 (new_sku, new_lines)。"""
        raise NotImplementedError


@dataclass
class SkuChangeRule(Rule):
    """
    ^ 操作符：根据条件修改SKU。
    格式：^条件 , 后缀
    示例：^[尺寸:L] , -1
    """
    condition: Condition
    suffix: str  # 如 "-1"

    def apply(self, sku: str, lines: List[str]) -> Tuple[str, List[str]]:
        if self.condition.exists_in_lines(lines):
            new_sku = _apply_sku_suffix(sku, self.suffix)
            logger.debug(f"  SKU变化: {sku} → {new_sku}（条件满足）")
            return new_sku, lines
        else:
            logger.debug(f"  SKU变化: 条件不满足，SKU保持 {sku}")
            return sku, lines


@dataclass
class DeleteRule(Rule):
    """
    ! 操作符：删除定制项中符合条件的整行。
    格式：!条件
    示例：![无:XYZ-NCT Larkspur]
    """
    condition: Condition

    def apply(self, sku: str, lines: List[str]) -> Tuple[str, List[str]]:
        new_lines = [ln for ln in lines if not self.condition.matches_line(ln)]
        removed = len(lines) - len(new_lines)
        if removed > 0:
            logger.debug(f"  删除规则: 移除了 {removed} 行")
        else:
            logger.debug(f"  删除规则: 未找到匹配行")
        return sku, new_lines


@dataclass
class TranslateRule(Rule):
    """
    = 操作符：翻译键名或键值对。
    格式：=源映射=目标映射
    示例：
        =[Name]=[名字]          → 键翻译（保留值）
        =[尺寸:L]=[尺寸:大号]    → 键值翻译
        =[尺寸:L|尺寸:M]=[尺寸:大号|尺寸:中号]  → 并列映射
    """
    mappings: List[Tuple[KeyValueCondition, KeyValueCondition]]
    # 每个元素是 (源条件, 目标条件)，其中目标条件的 value 是新值

    def apply(self, sku: str, lines: List[str]) -> Tuple[str, List[str]]:
        new_lines = []
        for line in lines:
            translated = False
            for src, tgt in self.mappings:
                if src.matches_line(line):
                    new_line = self._build_target_line(line, src, tgt)
                    new_lines.append(new_line)
                    translated = True
                    logger.debug(f"  翻译: {line!r} → {new_line!r}")
                    break  # 第一个匹配的映射生效
            if not translated:
                new_lines.append(line)
        return sku, new_lines

    @staticmethod
    def _build_target_line(line: str, src: KeyValueCondition, tgt: KeyValueCondition) -> str:
        """
        根据源条件和目标条件构造新行。

        情况1：目标有值 → 整体替换为目标键值对
        情况2：源键匹配 → 改键名，保留原值
        情况3：源值匹配 → 保留原键，改值
        """
        line = line.strip()
        if ":" in line:
            line_key, line_value = line.split(":", 1)
        else:
            line_key, line_value = line, ""

        if tgt.value is not None:
            # 目标指定了完整键值对 → 整体替换
            return f"{tgt.key}:{tgt.value}"

        # 目标无值：根据匹配类型决定替换键还是值
        mt = src.match_type(line)
        if mt == "value":
            # 源匹配的是值 → 保留原键，替换值
            return f"{line_key.strip()}:{tgt.key}"
        else:
            # 源匹配的是键（或 both）→ 替换键，保留原值
            return f"{tgt.key}:{line_value.strip()}" if line_value else tgt.key


@dataclass
class AppendRule(Rule):
    """
    + 操作符：在定制项末尾添加新行。
    格式：+[内容]
    示例：+[注意:加急]
    """
    content: str  # 要添加的整行内容（已恢复转义）

    def apply(self, sku: str, lines: List[str]) -> Tuple[str, List[str]]:
        logger.debug(f"  行末添加: {self.content!r}")
        return sku, lines + [self.content]


@dataclass
class InsertRule(Rule):
    """
    ++ 操作符：在某行下方添加新行（原行保留）。
    格式：[位置条件] ++ [内容]
    示例：[前面有云:白云] ++ [背后有云:白云]
    """
    position_condition: Condition  # 定位条件
    content: str  # 要插入的整行内容（已恢复转义）

    def apply(self, sku: str, lines: List[str]) -> Tuple[str, List[str]]:
        new_lines = []
        inserted = False
        for line in lines:
            new_lines.append(line)
            if self.position_condition.matches_line(line):
                new_lines.append(self.content)
                inserted = True
                logger.debug(f"  指定位置添加: 在 {line!r} 下方插入 {self.content!r}")
        if not inserted:
            logger.warning(f"  指定位置添加: 未找到匹配位置，内容 {self.content!r} 未插入")
        return sku, new_lines


# ──────────────────────────────────────────────
# 规则解析主函数
# ──────────────────────────────────────────────

def _find_unescaped(s: str, target: str) -> int:
    """在字符串中查找未转义的目标子串，返回索引；找不到返回 -1。"""
    # 此时字符串已经过 _escape_special 处理，特殊字符已是占位符
    return s.find(target)


def parse_rule(rule_str: str) -> List[Rule]:
    """
    解析单条规则字符串，返回 Rule 列表（; 分隔的子规则各产生一条 Rule）。

    执行顺序已在 Rule.apply 中按操作符类型排序，此处只负责解析。
    """
    if not rule_str or not isinstance(rule_str, str):
        return []

    original = rule_str.strip()
    if not original:
        return []

    # 第〇步：标准化换行符为 ;（Excel 单元格内可能使用换行或 <br> 分隔多条规则）
    original = original.replace("\r\n", ";").replace("\n", ";").replace("<br>", ";").replace("<BR>", ";")

    # 第一步：转义处理
    escaped = _escape_special(original)

    # 第二步：按 ; 分割子规则
    sub_rules = _split_by_operator(escaped, ";")
    if not sub_rules:
        return []

    rules: List[Rule] = []
    for sub in sub_rules:
        sub = sub.strip()
        if not sub:
            continue
        try:
            parsed = _parse_single_rule(sub, original)
            if parsed is not None:
                rules.append(parsed)
        except RuleParseError as e:
            logger.error(f"规则解析错误: {e}（原始规则: {original!r}）")
            continue
        except Exception as e:
            logger.error(f"规则解析异常: {e}（原始规则: {original!r}）")
            continue

    return rules


def _find_implicit_equals(s: str) -> int:
    """
    查找顶层 ]= 的位置（支持 [源]=[目标] 和 [源]=目标 两种格式）。
    返回 ] 的位置，找不到返回 -1。
    """
    depth = 0
    for i, ch in enumerate(s):
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0 and i + 1 < len(s) and s[i + 1] == "=":
                return i
    return -1


def _parse_single_rule(escaped: str, original: str) -> Optional[Rule]:
    """解析单条子规则（已转义处理）。"""
    s = escaped.strip()

    # ── 识别操作符 ──
    if s.startswith("^"):
        return _parse_sku_change(s, original)
    elif s.startswith("!"):
        return _parse_delete(s, original)
    elif s.startswith("="):
        return _parse_translate(s, original)
    elif s.startswith("[") and _find_implicit_equals(s) >= 0:
        # 隐式 = 规则：[源]=目标 或 [源]=[目标]（无前导 =）
        return _parse_translate("=" + s, original)
    elif _find_top_level_double_plus(s) >= 0:
        # 顶层（非 [] 内）有 ++ 才是插入规则
        return _parse_insert(s, original)
    elif s.startswith("+"):
        return _parse_append(s, original)
    else:
        raise RuleParseError(f"无法识别的操作符: {original!r}")


def _parse_sku_change(escaped: str, original: str) -> SkuChangeRule:
    """
    解析 ^ 规则：^条件 , 后缀

    找第一个 ] 后面的 ,（不在嵌套 [] 内）。
    """
    s = escaped[1:].strip()  # 去掉 ^
    if not s:
        raise RuleParseError(f"^ 规则缺少条件: {original!r}")

    # 找条件和后缀的分隔点：第一个 ] 后面的 ,
    comma_idx = _find_top_level_comma(s)
    if comma_idx < 0:
        raise RuleParseError(f"^ 规则缺少逗号分隔: {original!r}")

    cond_str = s[:comma_idx].strip()
    suffix_raw = s[comma_idx + 1:].strip()

    if not cond_str:
        raise RuleParseError(f"^ 规则条件为空: {original!r}")
    if not suffix_raw:
        raise RuleParseError(f"^ 规则后缀为空: {original!r}")

    condition = parse_condition(cond_str)
    # 后缀恢复转义
    suffix = _unescape(suffix_raw)

    return SkuChangeRule(operator="^", original=original, condition=condition, suffix=suffix)


def _find_top_level_comma(s: str) -> int:
    """找顶层（不在 [] 内）的逗号位置。"""
    depth = 0
    for i, ch in enumerate(s):
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
        elif ch == "," and depth == 0:
            return i
    return -1


def _parse_delete(escaped: str, original: str) -> DeleteRule:
    """解析 ! 规则：!条件"""
    s = escaped[1:].strip()  # 去掉 !
    if not s:
        raise RuleParseError(f"! 规则缺少条件: {original!r}")

    condition = parse_condition(s)
    return DeleteRule(operator="!", original=original, condition=condition)


def _parse_translate(escaped: str, original: str) -> TranslateRule:
    """
    解析 = 规则：=源=目标

    支持两种格式：
      =[源]=[目标]  — 目标有 [] 包裹（内部 | 分隔并列映射）
      =[源]=目标    — 目标无 [] 包裹（单一映射）
    """
    s = escaped[1:].strip()  # 去掉首个 =
    if not s:
        raise RuleParseError(f"= 规则内容为空: {original!r}")

    # 找顶层 ]= 分割点
    bracket_idx = _find_implicit_equals(s)
    if bracket_idx < 0:
        raise RuleParseError(f"= 规则格式错误（缺少 ]= 分隔）: {original!r}")

    src_part = s[:bracket_idx + 1].strip()   # 包括末尾的 ]
    tgt_part = s[bracket_idx + 2:].strip()   # 跳过 ]= 两个字符

    # 提取源 [...] 内容
    src_inner = _extract_brackets(src_part)
    if src_inner is None:
        raise RuleParseError(f"= 规则源格式错误: {original!r}")
    src_inner = src_inner.strip()
    if not src_inner:
        raise RuleParseError(f"= 规则源为空: {original!r}")

    # 提取目标内容：如果有 [] 则提取内部，否则整个作为单一目标
    if tgt_part.startswith("["):
        tgt_inner = _extract_brackets(tgt_part)
        if tgt_inner is None:
            raise RuleParseError(f"= 规则目标格式错误: {original!r}")
        tgt_inner = tgt_inner.strip()
    else:
        # 目标无 [] 包裹，整个作为单一键名
        tgt_inner = tgt_part.strip()
    if not tgt_inner:
        raise RuleParseError(f"= 规则目标为空: {original!r}")

    # 按 | 分割映射（但只分割顶层的 |）
    src_mappings = _split_top_level_pipe(src_inner)
    tgt_mappings = _split_top_level_pipe(tgt_inner)

    if len(src_mappings) != len(tgt_mappings):
        raise RuleParseError(
            f"= 规则源映射数量({len(src_mappings)})与目标映射数量({len(tgt_mappings)})不一致: {original!r}"
        )

    mappings = []
    for src_map, tgt_map in zip(src_mappings, tgt_mappings):
        src_cond = _parse_bracket_content(_unescape(src_map))
        tgt_cond = _parse_bracket_content(_unescape(tgt_map))
        mappings.append((src_cond, tgt_cond))

    return TranslateRule(operator="=", original=original, mappings=mappings)


def _split_top_level_pipe(s: str) -> List[str]:
    """按 | 分割，但忽略嵌套 [...] 内的 |。"""
    parts = []
    depth = 0
    current = []
    for ch in s:
        if ch == "[":
            depth += 1
            current.append(ch)
        elif ch == "]":
            depth -= 1
            current.append(ch)
        elif ch == "|" and depth == 0:
            parts.append("".join(current).strip())
            current = []
        else:
            current.append(ch)
    parts.append("".join(current).strip())
    return [p for p in parts if p]


def _parse_append(escaped: str, original: str) -> AppendRule:
    """解析 + 规则：+[内容]"""
    s = escaped[1:].strip()  # 去掉 +
    if not s:
        raise RuleParseError(f"+ 规则内容为空: {original!r}")

    inner = _extract_brackets(s)
    if inner is None:
        raise RuleParseError(f"+ 规则格式错误（需要 [...]) : {original!r}")
    inner = inner.strip()
    if not inner:
        raise RuleParseError(f"+ 规则内容为空: {original!r}")

    content = _unescape(inner)
    return AppendRule(operator="+", original=original, content=content)


def _parse_insert(escaped: str, original: str) -> InsertRule:
    """
    解析 ++ 规则：[位置条件] ++ [内容]
    """
    # 找顶层的 ++
    split_idx = _find_top_level_double_plus(escaped)
    if split_idx < 0:
        raise RuleParseError(f"++ 规则格式错误: {original!r}")

    pos_part = escaped[:split_idx].strip()
    content_part = escaped[split_idx + 2:].strip()

    if not pos_part or not content_part:
        raise RuleParseError(f"++ 规则内容不完整: {original!r}")

    position_condition = parse_condition(pos_part)

    inner = _extract_brackets(content_part)
    if inner is None:
        raise RuleParseError(f"++ 规则内容格式错误（需要 [...]）: {original!r}")
    inner = inner.strip()
    if not inner:
        raise RuleParseError(f"++ 规则内容为空: {original!r}")

    content = _unescape(inner)
    return InsertRule(operator="++", original=original, position_condition=position_condition, content=content)


def _find_top_level_double_plus(s: str) -> int:
    """找顶层（不在 [] 内）的 ++ 位置。"""
    depth = 0
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
        elif ch == "+" and depth == 0:
            if i + 1 < len(s) and s[i + 1] == "+":
                return i
        i += 1
    return -1


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第四部分：SKU后缀处理                                ║
# ╚════════════════════════════════════════════════════════════════════════════╝

def _apply_sku_suffix(sku: str, suffix: str) -> str:
    """
    在 SKU 上应用后缀。

    规则：
    - 如果 SKU 已有 -数字 后缀，则替换为新后缀
    - 如果 SKU 还没有 -数字 后缀，则追加后缀
    - 例如：
        ABC + -1    → ABC-1
        ABC-1 + -2  → ABC-2（替换，不是叠加）
        ABC-1 + -1  → ABC-1（相同后缀，不重复添加）
    """
    suffix = suffix.strip()
    if not suffix:
        return sku

    # 确保后缀以 - 开头
    if not suffix.startswith("-"):
        suffix = "-" + suffix

    # 检查 SKU 是否已有 -数字 后缀
    # 匹配末尾的 -数字（可能包含字母，如 -v2，但我们简单匹配 -后面的部分）
    match = re.search(r"-(.+)$", sku)
    if match:
        existing_suffix = "-" + match.group(1)
        if existing_suffix == suffix:
            # 相同后缀，不重复添加
            return sku
        else:
            # 替换旧后缀
            return sku[:match.start()] + suffix
    else:
        # 没有后缀，直接追加
        return sku + suffix


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第五部分：规则执行引擎                               ║
# ╚════════════════════════════════════════════════════════════════════════════╝

# 操作符执行顺序
_OPERATOR_ORDER = {"^": 0, "!": 1, "=": 2, "++": 3, "+": 4}


def normalize_customization(customization: str) -> List[str]:
    """
    统一将定制项字符串按 <br> 或 \\n 拆分为行列表。
    同时处理 \\r\\n、\\n 和 <br>（含大小写变体 <BR> <Br> <bR>）等情况，去除空行。
    每行去除首尾空白。
    """
    import re as _re
    if not customization:
        return []
    text = customization.replace("\r\n", "\n")
    # 统一将各种 <br> 变体替换为 \n
    text = _re.sub(r'<br\s*/?\s*>', '\n', text, flags=_re.IGNORECASE)
    lines = text.split("\n")
    return [ln.strip() for ln in lines if ln.strip()]


def apply_rules(rules: List[Rule], sku: str, customization: str) -> Tuple[str, str]:
    """
    对单条订单应用规则列表，返回 (new_sku, new_customization)。

    执行顺序：^ → ! → = → ++ → +
    """
    # 按操作符优先级排序
    sorted_rules = sorted(rules, key=lambda r: _OPERATOR_ORDER.get(r.operator, 99))

    # 统一分隔符处理
    lines = normalize_customization(customization)

    current_sku = sku
    current_lines = list(lines)

    for rule in sorted_rules:
        try:
            current_sku, current_lines = rule.apply(current_sku, current_lines)
        except Exception as e:
            logger.error(f"规则执行异常: {e}（规则: {rule.original!r}）")

    # 重新组合定制项（用 \n 换行，Excel 自动识别为单元格内多行）
    new_customization = "\n".join(current_lines)

    return current_sku, new_customization


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                第六.五部分：规则草稿生成（反推翻译模板）                    ║
# ╚════════════════════════════════════════════════════════════════════════════╝

def generate_a_rule_draft(before_text: str, after_text: str) -> str:
    """从「翻译前定制项」与「翻译后定制项」的差异，反推 A 系统翻译模板规则草稿。

    定位：系统生成 0.5，运营审核改到 1.0。覆盖机械部分：
      - 删除行   → ![键:值];               （before 有、after 无）
      - 键翻译   → =[旧键]=[新键];          （值不变、键改名）
      - 键值翻译 → =[旧键:旧值]=[新键:新值];（键或值变化）
      - 新增行   → +[整行];                （after 有、before 无，追加到末尾）
    每条规则以 ';' 结尾，对齐手写规范。

    设计取舍（草稿定位，非全自动）：
      - 仅做行级差异：按「键优先、其次值」匹配 before/after 各行。
      - 行间插入（++ 操作符）不自动生成：多出的行一律按「末尾追加(+)」表达，
        需人工判断是否应改为 ++ 定位插入。
      - SKU 变体（^ 操作符）与条件守卫不自动生成。
      - 同名键多行、键与值同时改变等复杂情形，退化为「删除+新增」，由人工合并。
    """

    def _kv(line: str):
        line = line.strip()
        if ":" in line:
            k, v = line.split(":", 1)
            return k.strip(), v.strip()
        return line, ""

    b_lines = [l.strip() for l in (before_text or "").splitlines() if l.strip()]
    a_lines = [l.strip() for l in (after_text or "").splitlines() if l.strip()]
    b = [(_kv(l), l) for l in b_lines]   # ((key, value), raw)
    a = [(_kv(l), l) for l in a_lines]

    a_consumed = [False] * len(a)
    deletes, translates, appends = [], [], []

    for (bk, bv), br in b:
        # 1) 同键匹配（键优先）
        mi = -1
        for i, ((ak, av), _ar) in enumerate(a):
            if not a_consumed[i] and ak == bk:
                mi = i
                break
        if mi >= 0:
            ak, av = a[mi][0]
            a_consumed[mi] = True
            if av == bv:
                continue  # 未变，无需规则
            translates.append(f"=[{bk}:{bv}]=[{ak}:{av}];")
            continue
        # 2) 同值匹配（键改名，值不变）→ 键翻译
        mi = -1
        if bv:
            for i, ((ak, av), _ar) in enumerate(a):
                if not a_consumed[i] and av == bv:
                    mi = i
                    break
        if mi >= 0:
            ak, _av = a[mi][0]
            a_consumed[mi] = True
            translates.append(f"=[{bk}]=[{ak}];")
            continue
        # 3) 删除
        deletes.append(f"![{bk}:{bv}];" if bv else f"![{br}];")

    # 4) 剩余 after 行 → 新增（追加末尾）
    for i, ((ak, av), ar) in enumerate(a):
        if not a_consumed[i]:
            appends.append(f"+[{ar}];")

    out = []
    if deletes:
        out.append("# 删除")
        out.extend(deletes)
    if translates:
        out.append("# 翻译")
        out.extend(translates)
    if appends:
        out.append("# 新增")
        out.extend(appends)
    if not out:
        return "# ✅ 前后定制项完全一致，无需生成规则"
    return "\n".join(out)


def strip_rule_comments(rule_text: str) -> str:
    """去除规则文本中以 # 开头的注释行（保留真实规则）。

    草稿生成器会在输出里加 # 删除 / # 翻译 / # 新增 分区注释方便人读，
    但 A 系统 parse_rule 不支持注释（会把换行当 ; 处理导致注释行报错）。
    在「校验草稿」与「写入翻译模板列」前调用本函数，确保喂给 parse_rule 的是干净规则。
    """
    if not rule_text:
        return ""
    kept = []
    for line in str(rule_text).splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        kept.append(line)
    return "\n".join(kept)


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第六部分：Excel处理                                  ║
# ╚════════════════════════════════════════════════════════════════════════════╝

def _find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    在DataFrame中查找列名（不区分大小写）。
    返回第一个匹配的列名，找不到返回 None。
    支持多行表头：匹配每列第一行文本（\\n 之前的部分）。
    """
    # 精确匹配优先
    for col in df.columns:
        for candidate in candidates:
            if col == candidate:
                return col
    # 大小写不敏感匹配
    for col in df.columns:
        for candidate in candidates:
            if col.lower() == candidate.lower():
                return col
    # 多行表头匹配：取列名第一行（\\n 之前）进行匹配
    for col in df.columns:
        col_first_line = str(col).split("\n")[0].strip()
        for candidate in candidates:
            if col_first_line == candidate:
                return col
            if col_first_line.lower() == candidate.lower():
                return col
    return None


def load_template(template_path: str) -> dict:
    """
    加载学习模板。

    返回：{SKU: 规则字符串} 的字典。
    如果同一SKU出现多次，使用最后出现的规则。
    """
    df = pd.read_excel(template_path, dtype=str)

    # 查找列（B列=第2列，但pandas用header读取，不一定按位置）
    # 尝试按列名查找
    sku_col = _find_column(df, ["SKU", "sku", "商品编码"])
    rule_col = _find_column(df, ["翻译模板", "规则", "rule", "template"])

    if sku_col is None:
        # 回退：B列是第2列（0-indexed: 1）
        if len(df.columns) >= 2:
            sku_col = df.columns[1]  # B列
        else:
            raise ValueError(f"无法识别模板文件中的SKU列。可用列: {list(df.columns)}")

    if rule_col is None:
        # 回退：C列是第3列（0-indexed: 2）
        if len(df.columns) >= 3:
            rule_col = df.columns[2]  # C列
        else:
            raise ValueError(f"无法识别模板文件中的规则列。可用列: {list(df.columns)}")

    logger.info(f"模板列识别: SKU={sku_col!r}, 规则={rule_col!r}")

    template: dict = {}
    for _, row in df.iterrows():
        sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
        rule_str = str(row[rule_col]).strip() if pd.notna(row[rule_col]) else ""
        if sku:
            # 同一SKU覆盖（保留最后出现的）
            template[sku] = rule_str

    # 预解析所有规则
    parsed_template: dict = {}
    parse_errors = 0
    for sku, rule_str in template.items():
        if rule_str:
            rules = parse_rule(rule_str)
            if rules:
                parsed_template[sku] = rules
            else:
                parse_errors += 1
                logger.warning(f"SKU {sku!r} 的规则解析失败，将跳过: {rule_str!r}")
        else:
            parsed_template[sku] = []

    if parse_errors:
        logger.warning(f"共有 {parse_errors} 条规则解析失败")
    logger.info(f"成功加载 {len(parsed_template)} 条SKU规则")

    return parsed_template


def load_orders(orders_path: str) -> pd.DataFrame:
    """加载原始订单。"""
    df = pd.read_excel(orders_path, dtype=str)
    logger.info(f"加载订单: {len(df)} 行, 列: {list(df.columns)}")
    return df


def process_orders(df: pd.DataFrame, template: dict) -> pd.DataFrame:
    """
    处理订单：对每一行应用规则。

    返回新的 DataFrame（包含修改后的 SKU 和定制项）。
    """
    # 识别列
    sku_col = _find_column(df, ["SKU", "sku", "商品编码"])
    cust_col = _find_column(df, ["定制项", "customization", "定制", "个性化"])

    if sku_col is None:
        raise ValueError(f"无法识别订单中的SKU列。可用列: {list(df.columns)}")
    if cust_col is None:
        raise ValueError(f"无法识别订单中的定制项列。可用列: {list(df.columns)}")

    logger.info(f"订单列识别: SKU={sku_col!r}, 定制项={cust_col!r}")

    result = df.copy()

    modified_count = 0
    skipped_count = 0
    error_count = 0

    for idx, row in result.iterrows():
        sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
        customization = str(row[cust_col]) if pd.notna(row[cust_col]) else ""

        if not sku:
            skipped_count += 1
            continue

        rules = template.get(sku)
        if rules is None or len(rules) == 0:
            # 未匹配到规则，跳过
            skipped_count += 1
            logger.debug(f"SKU {sku!r}: 无匹配规则，跳过")
            continue

        try:
            new_sku, new_cust = apply_rules(rules, sku, customization)
            result.at[idx, sku_col] = new_sku
            result.at[idx, cust_col] = new_cust
            modified_count += 1
            logger.debug(f"SKU {sku!r}: 处理完成")
        except Exception as e:
            logger.error(f"SKU {sku!r}: 处理异常: {e}")
            error_count += 1

    logger.info(f"处理完成: 修改 {modified_count} 行, 跳过 {skipped_count} 行, 错误 {error_count} 行")
    return result


def _col_idx_to_letter(idx: int) -> str:
    """1-based column index to Excel column letter(s). 1→A, 27→AA, 53→BA."""
    result = ''
    while idx > 0:
        idx -= 1
        result = chr(ord('A') + idx % 26) + result
        idx //= 26
    return result


def _col_ref_key(ref: str) -> tuple:
    """Sort key for cell references like 'A1', 'AC2', 'AY50'.
    Returns (col_letter_length, col_letter, row_number) so that
    'AC' sorts before 'AY' (both 2-letter, alphabetical order)."""
    col = ''.join(ch for ch in ref if ch.isalpha())
    row = int(''.join(ch for ch in ref if ch.isdigit())) if any(ch.isdigit() for ch in ref) else 0
    return (len(col), col, row)


def _update_xlsx_cells_lightweight(
    src_path: str,
    dst_path: str,
    updates_by_row: dict,
) -> None:
    """
    轻量级 XLSX 单元格更新——直接操作 ZIP/XML，跳过图片等二进制资源。

    完全保留原始格式、样式、列宽、合并单元格、图片等。
    对含大量图片的订单文件友好：不解析图片，不依赖 PIL/Pillow。

    Args:
        src_path: 源 XLSX 文件路径
        dst_path: 目标 XLSX 文件路径
        updates_by_row: {excel_row_1based: {col_letter: new_value}}
    """
    import zipfile
    import xml.etree.ElementTree as ET
    import shutil
    import tempfile
    import os

    MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    def _tag(local: str) -> str:
        return f'{{{MAIN_NS}}}{local}'

    # 注册命名空间以保持输出整洁（避免 ns0: 前缀）
    ET.register_namespace('', MAIN_NS)
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    modified_parts: dict = {}   # {arcname_in_zip: serialized_bytes}
    sst_entries: list = []       # shared strings table entries
    sst_modified = False

    # ── 第一步：读取并解析需要修改的 XML ──
    with zipfile.ZipFile(src_path, 'r') as zf:
        # 1a. 确定活动工作表文件路径
        sheet_path = 'xl/worksheets/sheet1.xml'
        rels_path = 'xl/_rels/workbook.xml.rels'

        if 'xl/workbook.xml' in zf.namelist():
            wb_xml = zf.read('xl/workbook.xml')
            wb_root = ET.fromstring(wb_xml)
            sheets_elem = wb_root.find(_tag('sheets'))
            if sheets_elem is not None:
                first_sheet = sheets_elem.find(_tag('sheet'))
                if first_sheet is not None:
                    r_id = first_sheet.get(
                        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
                    )
                    if r_id and rels_path in zf.namelist():
                        rels_xml = zf.read(rels_path)
                        rels_root = ET.fromstring(rels_xml)
                        for rel in rels_root:
                            if rel.get('Id') == r_id:
                                target = rel.get('Target', '')
                                if target:
                                    sheet_path = f'xl/{target}' if not target.startswith('/') else target
                                break

        # 1b. 读取共享字符串表 (SharedStrings)
        sst_path = 'xl/sharedStrings.xml'
        sst_exists = sst_path in zf.namelist()
        if sst_exists:
            sst_xml = zf.read(sst_path)
            sst_root = ET.fromstring(sst_xml)
            for si in sst_root.findall(_tag('si')):
                t_elem = si.find(_tag('t'))
                if t_elem is not None:
                    sst_entries.append(t_elem.text or '')
                    continue
                # 富文本 <r><t>...</t></r>
                text_parts = []
                for r_elem in si.findall(_tag('r')):
                    rt = r_elem.find(_tag('t'))
                    if rt is not None:
                        text_parts.append(rt.text or '')
                sst_entries.append(''.join(text_parts))
            logger.debug(f"共享字符串表: {len(sst_entries)} 条")

        # 1c. 读取工作表 XML
        if sheet_path not in zf.namelist():
            raise FileNotFoundError(f"工作表文件不存在: {sheet_path}")
        ws_xml = zf.read(sheet_path)
        ws_root = ET.fromstring(ws_xml)

        sheetdata = ws_root.find(_tag('sheetData'))
        if sheetdata is None:
            raise ValueError("工作表中无 sheetData 元素")

        # 建立已有行索引：row_num → row_element
        row_map = {}
        for row_elem in sheetdata.findall(_tag('row')):
            r = row_elem.get('r')
            if r:
                row_map[int(r)] = row_elem

        # 1d. 按行修改目标单元格
        for row_num, col_updates in sorted(updates_by_row.items()):
            # 获取或创建 <row>
            if row_num in row_map:
                row_elem = row_map[row_num]
            else:
                row_elem = ET.SubElement(sheetdata, _tag('row'))
                row_elem.set('r', str(row_num))
                row_map[row_num] = row_elem

            # 建立该行的列索引：col_letter → c_element
            cell_map = {}
            for c_elem in row_elem.findall(_tag('c')):
                cr = c_elem.get('r', '')
                col_letter = ''.join(ch for ch in cr if ch.isalpha())
                if col_letter:
                    cell_map[col_letter] = c_elem

            for col_letter, new_value in col_updates.items():
                new_value = str(new_value) if new_value is not None else ''

                if col_letter in cell_map:
                    c_elem = cell_map[col_letter]
                    t_attr = c_elem.get('t', '')
                else:
                    # 新建单元格
                    c_elem = ET.SubElement(row_elem, _tag('c'))
                    c_elem.set('r', f'{col_letter}{row_num}')
                    t_attr = ''
                    cell_map[col_letter] = c_elem

                if t_attr == 's':
                    # 共享字符串：追加到 SST，更新 <v> 索引
                    sst_entries.append(new_value)
                    new_idx = len(sst_entries) - 1
                    sst_modified = True

                    v_elem = c_elem.find(_tag('v'))
                    if v_elem is None:
                        # 清理旧的内联字符串子元素
                        for child in list(c_elem):
                            if child.tag == _tag('is'):
                                c_elem.remove(child)
                        v_elem = ET.SubElement(c_elem, _tag('v'))
                    v_elem.text = str(new_idx)
                    c_elem.set('t', 's')
                elif t_attr == 'inlineStr':
                    # 内联字符串：直接修改 <is><t> 文本
                    is_elem = c_elem.find(_tag('is'))
                    if is_elem is None:
                        for child in list(c_elem):
                            if child.tag in (_tag('v'), _tag('f')):
                                c_elem.remove(child)
                        is_elem = ET.SubElement(c_elem, _tag('is'))
                    t_elem = is_elem.find(_tag('t'))
                    if t_elem is None:
                        t_elem = ET.SubElement(is_elem, _tag('t'))
                    t_elem.text = new_value
                    t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    c_elem.set('t', 'inlineStr')
                else:
                    # 无 t 属性或 t="n"/"str" → 转为内联字符串
                    is_elem = c_elem.find(_tag('is'))
                    if is_elem is None:
                        for child in list(c_elem):
                            if child.tag in (_tag('v'), _tag('f')):
                                c_elem.remove(child)
                        is_elem = ET.SubElement(c_elem, _tag('is'))
                    t_elem = is_elem.find(_tag('t'))
                    if t_elem is None:
                        t_elem = ET.SubElement(is_elem, _tag('t'))
                    t_elem.text = new_value
                    t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    c_elem.set('t', 'inlineStr')

        # 序列化修改后的工作表 XML
        modified_parts[sheet_path] = ET.tostring(
            ws_root, xml_declaration=True, encoding='UTF-8'
        )

        # 序列化修改后的 SST（如果有修改）
        if sst_modified and sst_exists:
            sst_root.set('count', str(len(sst_entries)))
            sst_root.set('uniqueCount', str(len(sst_entries)))
            modified_parts[sst_path] = ET.tostring(
                sst_root, xml_declaration=True, encoding='UTF-8'
            )

    # ── 第二步：流式写回 ZIP（图片等二进制资源原样复制）──
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    tmp.close()
    try:
        with zipfile.ZipFile(src_path, 'r') as zin:
            with zipfile.ZipFile(tmp.name, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in modified_parts:
                        zout.writestr(item, modified_parts[item.filename])
                    else:
                        zout.writestr(item, zin.read(item.filename))
        shutil.move(tmp.name, dst_path)
    finally:
        if os.path.exists(tmp.name):
            os.unlink(tmp.name)
    logger.debug(f"轻量级 XLSX 更新完成: 修改了 {len(modified_parts)} 个 ZIP 部件")


def process_orders_preserve_format(orders_path, output_path, template):
    """
    用 zipfile 直接操作 XLSX 底层 ZIP，只修改目标单元格 XML，保留一切原始内容。
    图片（xl/media/）、格式（xl/styles.xml）、列宽、行高、公式、合并单元格完全不动。

    流程：pandas 读取数据 → 处理 → 解压到临时目录 → 修改 sheet1.xml → 重新打包。

    返回：(modified_count, skipped_count, error_count)
    """
    import zipfile
    import xml.etree.ElementTree as ET
    import tempfile
    import os
    import shutil

    # ── 第1步：pandas 读取全部数据，识别列名 ──
    df = pd.read_excel(orders_path, dtype=str)
    sku_col_name = _find_column(df, ["SKU", "sku", "商品编码"])
    cust_col_name = _find_column(df, ["定制项", "customization", "定制", "个性化"])

    if sku_col_name is None:
        if len(df.columns) >= 5:
            sku_col_name = df.columns[4]
        else:
            raise ValueError(f"无法识别订单中的SKU列。可用列: {list(df.columns)}")
    if cust_col_name is None:
        if len(df.columns) >= 9:
            cust_col_name = df.columns[8]
        else:
            raise ValueError(f"无法识别订单中的定制项列。可用列: {list(df.columns)}")

    logger.info(f"订单列识别: SKU={sku_col_name!r}, 定制项={cust_col_name!r}")

    logger.info(f"读取 {len(df)} 行 × {len(df.columns)} 列")

    # 计算列字母（用于 XML 单元格引用）
    all_cols = list(df.columns)
    sku_col_1based = all_cols.index(sku_col_name) + 1
    cust_col_1based = all_cols.index(cust_col_name) + 1
    sku_col_letter = _col_idx_to_letter(sku_col_1based)
    cust_col_letter = _col_idx_to_letter(cust_col_1based)

    # ── 识别加急列（K列：按列名优先，回退到固定位置索引10）──
    urgent_col_name = _find_column(df, ["加急", "urgent", "紧急", "加急标志"])
    if urgent_col_name:
        urgent_col_index = list(df.columns).index(urgent_col_name)
        logger.info(f"加急列识别: {urgent_col_name!r} (索引 {urgent_col_index})")
    else:
        urgent_col_index = 10  # 回退到固定 K 列
        logger.info(f"加急列识别: 未找到列名，使用固定索引 {urgent_col_index}")

    # ── AC/AS 列选择：优先按列名匹配，找不到再用固定列号 ──
    channel_col_name = _find_column(df, ["物流渠道", "运输方式", "物流", "渠道", "买家选择渠道类型"])
    delivery_col_name = _find_column(df, ["时效", "配送时效", "时效类型", "购买配送服务"])

    if channel_col_name:
        channel_col_letter = _col_idx_to_letter(list(df.columns).index(channel_col_name) + 1)
        logger.info(f"物流渠道列: {channel_col_name!r} -> {channel_col_letter}")
    else:
        channel_col_letter = 'AC'
        logger.info(f"物流渠道列: 固定位置 {channel_col_letter}")

    if delivery_col_name:
        delivery_col_letter = _col_idx_to_letter(list(df.columns).index(delivery_col_name) + 1)
        logger.info(f"时效列: {delivery_col_name!r} -> {delivery_col_letter}")
    else:
        delivery_col_letter = 'AS'
        logger.info(f"时效列: 固定位置 {delivery_col_letter}")

    logger.info(f"DataFrame: {len(df)} 行, {len(df.columns)} 列")

    # ── 第2步：遍历处理，收集变更 ──
    modified_count = 0
    skipped_count = 0
    error_count = 0
    updates = {}  # {excel_row_1based: {col_letter: new_value}}

    ac_as_filled = 0  # 统计 AC/AS 填充次数
    ac_as_map = {}     # {excel_row: {col_letter: value}} — 供 openpyxl 覆写
    for idx, row in df.iterrows():
        excel_row = idx + 2  # +2: 0-based pandas index → 1-based Excel row

        sku = str(row[sku_col_name]).strip() if pd.notna(row[sku_col_name]) else ""

        if not sku:
            skipped_count += 1
            continue

        # ── 确保此行在 updates 中有条目录入 ──
        row_updates = updates.setdefault(excel_row, {})

        # ── AC/AS 自动填充（按列名或固定位置读取加急列）──
        if channel_col_letter and delivery_col_letter:
            urgent_val = ""
            if urgent_col_index is not None and urgent_col_index < len(df.columns):
                raw = row.iloc[urgent_col_index]
                urgent_val = str(raw).strip().upper() if pd.notna(raw) else ""
            ch_val = "快递" if urgent_val == "Y" else "经济线"
            dv_val = "2" if urgent_val == "Y" else "1"
            row_updates[channel_col_letter] = ch_val
            row_updates[delivery_col_letter] = dv_val
            ac_as_filled += 1
            ac_as_map.setdefault(excel_row, {})[channel_col_letter] = ch_val
            ac_as_map.setdefault(excel_row, {})[delivery_col_letter] = dv_val

        # ── 翻译逻辑（只对匹配规则的行执行） ──
        customization = str(row[cust_col_name]) if pd.notna(row[cust_col_name]) else ""

        rules = template.get(sku)
        if not rules:
            skipped_count += 1
        else:
            try:
                new_sku, new_cust = apply_rules(rules, sku, customization)
                if new_sku != sku or new_cust != customization:
                    row_updates[sku_col_letter] = new_sku
                    row_updates[cust_col_letter] = new_cust
                    modified_count += 1
            except Exception as e:
                logger.error(f"SKU {sku!r}: 规则执行异常: {e}")
                error_count += 1

    logger.info(f"AC/AS填充完成: 共填充 {ac_as_filled} 行, updates={len(updates)} 行")

    # ── 第3步：解压 → 修改 sheet1.xml → 重新打包 ──
    if updates:
        tmpdir = tempfile.mkdtemp()
        try:
            # 3a. 解压整个 XLSX 到临时目录（图片、格式等全部原样提取）
            with zipfile.ZipFile(orders_path, 'r') as zf:
                zf.extractall(tmpdir)

            # 3b. 找到工作表文件
            sheet_path = os.path.join(tmpdir, 'xl', 'worksheets', 'sheet1.xml')
            if not os.path.exists(sheet_path):
                raise FileNotFoundError(f"工作表文件不存在: {sheet_path}")

            # 修复1：先读取原始 XML，清理重复的 xmlns 属性后再解析
            _ns_uri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            _r_uri  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            with open(sheet_path, 'r', encoding='utf-8') as _f:
                _raw_xml = _f.read()
            # 移除根元素上所有 xmlns 声明（ElementTree 解析后 register_namespace 会重新生成）
            _raw_xml = re.sub(
                r'\s+xmlns(?::\w+)?\s*=\s*"[^"]*"',
                '',
                _raw_xml,
                count=3,
            )
            # 在根元素上添加干净的命名空间声明
            _raw_xml = _raw_xml.replace(
                '<worksheet',
                f'<worksheet xmlns="{_ns_uri}" xmlns:r="{_r_uri}"',
                1,
            )
            with open(sheet_path, 'w', encoding='utf-8') as _f:
                _f.write(_raw_xml)

            # 注册命名空间（告诉 ElementTree 序列化时使用这些前缀）
            ET.register_namespace('', _ns_uri)
            ET.register_namespace('r', _r_uri)

            tree = ET.parse(sheet_path)
            root = tree.getroot()

            ns = '{%s}' % _ns_uri

            sheetdata = root.find(f'{ns}sheetData')
            if sheetdata is None:
                raise ValueError("工作表中无 sheetData 元素")

            # 建立已有行的索引
            row_map = {}
            for row_elem in sheetdata.findall(f'{ns}row'):
                r = row_elem.get('r')
                if r:
                    row_map[int(r)] = row_elem

            # 3c. 准备 sharedStrings：读取已有字符串，为新值分配索引
            shared_path = os.path.join(tmpdir, 'xl', 'sharedStrings.xml')
            shared_strings = []
            shared_count_attr = '0'
            if os.path.exists(shared_path):
                _ss_tree = ET.parse(shared_path)
                _ss_root = _ss_tree.getroot()
                shared_count_attr = _ss_root.get('count', '0')
                for _si in _ss_root.findall(f'{ns}si'):
                    # 优先读取 <t> 简单文本；若无则处理 <r> 富文本
                    _t_elem = _si.find(f'{ns}t')
                    if _t_elem is not None and _t_elem.text is not None:
                        shared_strings.append(_t_elem.text)
                    else:
                        # 富文本：提取所有 <r><t> 的文本拼接
                        _r_elems = _si.findall(f'{ns}r')
                        if _r_elems:
                            _parts = []
                            for _r in _r_elems:
                                _rt = _r.find(f'{ns}t')
                                if _rt is not None and _rt.text:
                                    _parts.append(_rt.text)
                            shared_strings.append(''.join(_parts))
                        else:
                            shared_strings.append('')
                logger.debug(f"sharedStrings 现有 {len(shared_strings)} 条")

            # 收集所有需要写入的新值，去重后分配索引
            _all_vals = []
            for _row_updates in updates.values():
                _all_vals.extend(_row_updates.values())
            _unique_vals = list(dict.fromkeys(str(v) for v in _all_vals))  # 保序去重
            _val_to_idx = {}
            for _v in _unique_vals:
                if _v not in shared_strings:
                    _val_to_idx[_v] = len(shared_strings)
                    shared_strings.append(_v)
                else:
                    _val_to_idx[_v] = shared_strings.index(_v)
            logger.debug(f"sharedStrings 新增 {len(_val_to_idx)} 条，总计 {len(shared_strings)} 条")

            # 3c-bis. 在 styles.xml 中注入宋体10号居中样式（供 AC/AS 单元格引用）
            _styles_path = os.path.join(tmpdir, 'xl', 'styles.xml')
            _ac_as_style_id = None
            if os.path.exists(_styles_path):
                try:
                    _st_tree = ET.parse(_styles_path)
                    _st_root = _st_tree.getroot()
                    _target_font_id = None
                    _fonts_elem = _st_root.find(f'{ns}fonts')
                    if _fonts_elem is not None:
                        for _fi, _font_elem in enumerate(_fonts_elem.findall(f'{ns}font')):
                            _name_elem = _font_elem.find(f'{ns}name')
                            _sz_elem = _font_elem.find(f'{ns}sz')
                            if (_name_elem is not None and _sz_elem is not None and
                                _name_elem.get('val', '') == '宋体' and _sz_elem.get('val', '') == '10'):
                                _target_font_id = _fi
                                break
                    if _target_font_id is None:
                        _target_font_id = 3
                    _cxfs = _st_root.find(f'{ns}cellXfs')
                    if _cxfs is not None:
                        _count = int(_cxfs.get('count', '0'))
                        _new_xf = ET.SubElement(_cxfs, f'{ns}xf')
                        _new_xf.set('numFmtId', '0')
                        _new_xf.set('fontId', str(_target_font_id))
                        _new_xf.set('fillId', '0')
                        _new_xf.set('borderId', '0')
                        _new_xf.set('xfId', '0')
                        _new_xf.set('applyFont', '1')
                        _new_xf.set('applyAlignment', '1')
                        _align_elem = ET.SubElement(_new_xf, f'{ns}alignment')
                        _align_elem.set('horizontal', 'center')
                        _align_elem.set('vertical', 'center')
                        _cxfs.set('count', str(_count + 1))
                        _ac_as_style_id = _count
                        _st_tree.write(_styles_path, xml_declaration=True, encoding='UTF-8')
                        logger.debug(f"styles.xml 注入样式 s={_ac_as_style_id} (fontId={_target_font_id})")
                except Exception as _se:
                    logger.warning(f"styles.xml 注入失败（将不设置格式）: {_se}")

            # 3d. 更新目标单元格（使用 sharedStrings 引用，兼容所有 Excel 版本）
            for excel_row, col_updates in sorted(updates.items()):
                # 获取或创建行元素
                if excel_row in row_map:
                    row_elem = row_map[excel_row]
                else:
                    row_elem = ET.SubElement(sheetdata, f'{ns}row')
                    row_elem.set('r', str(excel_row))
                    row_map[excel_row] = row_elem

                # 建立该行已有单元格的列字母索引
                cell_map = {}
                for c in row_elem.findall(f'{ns}c'):
                    cr = c.get('r', '')
                    cl = ''.join(ch for ch in cr if ch.isalpha())
                    if cl:
                        cell_map[cl] = c

                for col_letter, new_val in col_updates.items():
                    new_val = str(new_val) if new_val is not None else ''
                    cell_ref = f'{col_letter}{excel_row}'
                    ss_idx = _val_to_idx[new_val]  # shared string 索引

                    if col_letter in cell_map:
                        c_elem = cell_map[col_letter]
                    else:
                        c_elem = ET.SubElement(row_elem, f'{ns}c')
                        c_elem.set('r', cell_ref)

                    # 清除旧的子元素（v / f / is）
                    for child in list(c_elem):
                        if child.tag in (f'{ns}v', f'{ns}f', f'{ns}is'):
                            c_elem.remove(child)

                    # 写入 shared string 引用（标准 OOXML 方式）
                    c_elem.set('t', 's')
                    # AC/AS 列设置居中宋体10号样式
                    if (_ac_as_style_id is not None and
                        col_letter in (channel_col_letter, delivery_col_letter)):
                        c_elem.set('s', str(_ac_as_style_id))
                    v_elem = ET.SubElement(c_elem, f'{ns}v')
                    v_elem.text = str(ss_idx)

                # 修复5：将本行所有 <c> 元素按列字母排序（Excel 要求升序）
                _cells = [(c.get('r', ''), c) for c in row_elem.findall(f'{ns}c')]
                _cells.sort(key=lambda x: _col_ref_key(x[0]))
                for _i, (_ref, _c) in enumerate(_cells):
                    row_elem.remove(_c)
                    row_elem.insert(_i, _c)

            # 写回修改后的 sheet1.xml
            tree.write(sheet_path, xml_declaration=True, encoding='UTF-8')

            # 3e. 写回 sharedStrings.xml（重建，避免 ET 命名空间问题）
            if _val_to_idx:
                _ss_lines = [
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    f'<sst xmlns="{_ns_uri}" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
                ]
                for _si_text in shared_strings:
                    _escaped = _si_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                    _ss_lines.append(f'<si><t xml:space="preserve">{_escaped}</t></si>')
                _ss_lines.append('</sst>')
                with open(shared_path, 'w', encoding='utf-8') as _f:
                    _f.write('\n'.join(_ss_lines))
                logger.debug(f"sharedStrings.xml 已重建: {len(shared_strings)} 条")

            # 修复2：验证文件写入成功
            if os.path.exists(sheet_path) and os.path.getsize(sheet_path) > 0:
                logger.debug(f"sheet1.xml 写入成功，大小: {os.path.getsize(sheet_path)} 字节")
            else:
                raise RuntimeError("sheet1.xml 写入失败")

            # 修复3：更新 [Content_Types].xml
            ct_path = os.path.join(tmpdir, '[Content_Types].xml')
            if os.path.exists(ct_path):
                ct_tree = ET.parse(ct_path)
                ct_root = ct_tree.getroot()
                
                found = False
                for elem in ct_root.findall('Override'):
                    if elem.get('PartName') == '/xl/worksheets/sheet1.xml':
                        found = True
                        break
                
                if not found:
                    override = ET.SubElement(ct_root, 'Override')
                    override.set('PartName', '/xl/worksheets/sheet1.xml')
                    override.set('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
                    ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')
                    logger.debug("已更新 [Content_Types].xml")

            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for dirpath, _, filenames in os.walk(tmpdir):
                    for fn in filenames:
                        full = os.path.join(dirpath, fn)
                        arc = os.path.relpath(full, tmpdir).replace('\\', '/')
                        zout.write(full, arc)

            total_cells = sum(len(v) for v in updates.values())
            logger.info(
                f"ZIP/XML 更新了 {len(updates)} 行（{total_cells} 个单元格）"
            )
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
    else:
        # 无变更，直接复制原文件
        shutil.copy2(orders_path, output_path)
        logger.info("无变更，直接复制原文件")

    # ── 第4步（已禁用）：不再使用 openpyxl 覆写，纯 XML sharedStrings 方式写入 AC/AS ──
    # AC/AS 列已在上面的 XML 修改步骤中通过 sharedStrings 写入，无需额外处理。
    # 如需恢复 openpyxl 方式，取消下方注释即可。
    # if ac_as_map and os.path.exists(output_path):
    #     ... (v2.3.x 的 openpyxl 覆写代码)

    logger.info(
        f"处理完成: 修改 {modified_count} 行, 跳过 {skipped_count} 行, 错误 {error_count} 行, "
        f"AC/AS填充 {ac_as_filled} 行"
    )
    return modified_count, skipped_count, error_count, ac_as_filled


# ╔════════════════════════════════════════════════════════════════════════════╗
# ║                       第七部分：命令行接口                                 ║
# ╚════════════════════════════════════════════════════════════════════════════╝

def main():
    parser = argparse.ArgumentParser(
        description="表格翻译系统 - 基于自定义语法的订单定制项翻译工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
    python translator.py --template template.xlsx --orders orders.xlsx --output result.xlsx
    python translator.py -t template.xlsx -o orders.xlsx -r result.xlsx

规则语法参考：
    ^[条件] , 后缀          SKU变化（根据条件修改SKU）
    ![条件]                删除定制项中符合条件的行
    =[源]=[目标]           翻译键名或键值对
    [位置条件] ++ [内容]    在指定位置下方添加新行
    +[内容]                在定制项末尾添加新行

内部修饰符：
    [键:值]  范围限定    &  逻辑与    |  逻辑或    ;  子规则分隔
    优先级：[] > & > |
    执行顺序：^ → ! → = → ++ → +
        """,
    )
    parser.add_argument(
        "--template", "-t",
        required=True,
        help="学习模板文件路径（Excel，包含SKU和翻译模板列）",
    )
    parser.add_argument(
        "--orders", "-o",
        required=True,
        help="原始订单文件路径（Excel，包含SKU和定制项列）",
    )
    parser.add_argument(
        "--output", "-r",
        default="result.xlsx",
        help="输出文件路径（默认: result.xlsx）",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="显示详细日志（DEBUG级别）",
    )

    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)

    logger.info("=" * 60)
    logger.info("表格翻译系统启动")
    logger.info(f"学习模板: {args.template}")
    logger.info(f"原始订单: {args.orders}")
    logger.info(f"输出文件: {args.output}")
    logger.info("=" * 60)

    # 1. 加载学习模板
    logger.info("加载学习模板...")
    try:
        template = load_template(args.template)
    except Exception as e:
        logger.error(f"加载学习模板失败: {e}")
        sys.exit(1)

    # 2. 加载原始订单
    logger.info("加载原始订单...")
    try:
        orders_df = load_orders(args.orders)
    except Exception as e:
        logger.error(f"加载原始订单失败: {e}")
        sys.exit(1)

    # 3. 处理订单（openpyxl 保留原始格式）
    logger.info("开始处理订单（保留原始格式）...")
    try:
        modified, skipped, errors, ac_as = process_orders_preserve_format(
            args.orders, args.output, template
        )
    except Exception as e:
        logger.error(f"处理订单失败: {e}")
        sys.exit(1)

    # 4. 输出结果
    logger.info(f"✓ 结果已保存到 {args.output}")
    logger.info(f"  修改 {modified} 行, 跳过 {skipped} 行, 错误 {errors} 行, AC/AS {ac_as} 行")

    logger.info("=" * 60)
    logger.info("处理完成!")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
