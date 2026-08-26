# -*- coding: utf-8 -*-
"""
build_callie_sku.py  —— 把属性表「入库」（通用，不动核心引擎）

用法：
  python build_callie_sku.py <SKU> <attribute.xlsx>

动作（方案 A —— 去冗余，属性表入库即止）：
  把 attribute.xlsx 编译成 callie_sku_configs/<SKU>/attribute_config.json 入库。
  规则不再手写 rule.json，统一在「B 系统导出」时由翻译模板的 D/E/F 列编译生成，
  且入库的属性表会被 B 系统自动复用（无需每次导出重复上传）。

核心引擎（callie_generator.py）完全不被修改 —— 换 SKU 只是换属性表文件。
"""

import os
import sys
import json

# 让脚本可直接运行（同目录的 callie_generator 能被 import）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import callie_generator as cg


def _norm(s):
    """归一化：去空格、转小写，便于跨命名风格（Hair Color / HairColour）匹配。"""
    return "".join(ch for ch in str(s).lower() if ch.isalnum())


def _detect(names, *needles):
    """通用子串匹配（兼容 空格/驼峰/美英拼写）。返回第一个命中的组名；无则 None。"""
    norm_names = [(n, _norm(n)) for n in names]
    for nd in needles:
        ndn = _norm(nd)
        for n, nn in norm_names:
            if ndn in nn:
                return n
    return None


def _detect_global(names, *needles):
    """全局字段（如 Style / Character）：在全部命中里，优先【不含 hair/beard 修饰】
    且【归一化长度最短】的组名（角色作用域组通常带前缀、更长，会被排在后面）。"""
    cands = []
    for n in names:
        nn = _norm(n)
        if any(nd in nn for nd in needles):
            cands.append(n)
    if not cands:
        return None
    cands.sort(key=lambda n: (("hair" in _norm(n)) or ("beard" in _norm(n)),
                              len(_norm(n))))
    return cands[0]


def _detect_base(names, *needles):
    """角色作用域字段（如 Hair Color / Hair Style / Beard Color / Beard Style）：
    取首个命中组，剥掉其 '{Role}'s ' 前缀，得到基础字段名。
    例：'Girl's Hair Color' -> 'Hair Color'；'Man's Beard Style' -> 'Beard Style'。
    若命中组本身已无前缀（如直接叫 'Hair Color'），则返回原样。"""
    hit = _detect(names, *needles)
    if not hit:
        return None
    # 剥前缀：常见分隔符 's / - / :
    for sep in ("'s ", " - ", ": ", " -"):
        if sep in hit:
            return hit.split(sep, 1)[-1].strip()
    return hit


def _draft_v2(attr, sku):
    """【已弃用】方案 A 后规则统一由翻译模板 D/E/F 编译，不再手写 rule.json 草稿。
    保留函数仅作历史参考，不再被调用。"""
    raise NotImplementedError("方案 A：不再生成朴素 rule.json，规则由翻译模板编译")


def main(sku=None, xlsx=None, rule_in=None):
    """
    两种调用方式：
      1) CLI：python build_callie_sku.py <SKU> <attribute.xlsx>
      2) 程序化：main(sku, xlsx_path) —— 本地版 Streamlit「属性表入库」用此方式。
    动作：把属性表编译成 callie_sku_configs/<SKU>/attribute_config.json 入库。
    rule_in 参数已废弃（方案 A 不再用 rule.json）。
    """
    if sku is None:  # CLI 模式：从 sys.argv 取参数
        if len(sys.argv) < 3:
            print("用法: python build_callie_sku.py <SKU> <attribute.xlsx>")
            sys.exit(1)
        sku = sys.argv[1]
        xlsx = sys.argv[2]
    if not sku or not xlsx:
        print("用法: python build_callie_sku.py <SKU> <attribute.xlsx>")
        sys.exit(1)

    base = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "callie_sku_configs", sku)
    os.makedirs(base, exist_ok=True)

    # 编译属性表 → attribute_config.json（入库；B 系统导出时自动复用，无需重复上传）
    attr_json = os.path.join(base, "attribute_config.json")
    attr = cg.save_attribute_config(xlsx, attr_json)
    n_opts = sum(1 for g in attr["groups"] if g.get("opts"))
    print(f"[ok] 属性表已入库 -> {attr_json}  ({len(attr['groups'])} 组，含选项 {n_opts} 组)")
    print("      B 系统导出时会按订单 SKU 自动复用此属性表；规则由翻译模板 D/E/F 编译。")
    return attr


if __name__ == "__main__":
    main()
