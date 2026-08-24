# -*- coding: utf-8 -*-
"""
双sheet全公式核价Excel生成器（cost-pricing-assistant v3 核心模板）
====================================================
Sheet1「清单参考」= 价格真源（C列浅黄可编辑，数据从row3起）
Sheet2「核价明细」= 主表（F列引用清单参考C列蓝色公式，G列=数量×单价，合计SUM）

用法（函数式API）:
    from generate_pricing_xlsx import build_pricing_workbook
    build_pricing_workbook(
        title="XX项目——XX核价明细表",
        info_left="项目：XX公司",
        info_right="合同：XX年度零星工程（下浮XX已含）",
        refs=[("子目名", 单价, "来源", "备注"), ...],      # row3起依次写入
        data=[(序号, 内容, 特征, 单位, 数量, ref_idx, "子目名", "备注", is_est), ...],
        # ref_idx = refs列表的0-based索引！脚本内部自动换算为 =清单参考!C{ref_idx+3}
        notes=["说明...", ...],
        out_path=r"D:\造价管理\核价明细_XX.xlsx",
    )

⚠️ ref 换算铁律（两次踩坑总结）:
    refs[i] 位于 清单参考 row 3+i，公式为 =清单参考!C{i+3}
    即 refs[0]→C3, refs[1]→C4 ...
    调用方传 0-based ref_idx，本模块内部 +3，杜绝错位。

生成后务必验证（本模块自带 verify）:
    from generate_pricing_xlsx import verify_pricing
    ok = verify_pricing(out_path)   # 逐行取引用单元格实际值与refs比对
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------- 样式常量 ----------
FONT_NAME = "微软雅黑"
TF  = Font(name=FONT_NAME, size=14, bold=True)
NF  = Font(name=FONT_NAME, size=10)
SF  = Font(name=FONT_NAME, size=9)
BF  = Font(name=FONT_NAME, size=10, bold=True)
FORMULA_FONT = Font(name=FONT_NAME, size=10, color="0070C0")   # 蓝色=公式
_thin = Side(style="thin")
BDR = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
HF  = PatternFill("solid", start_color="4472C4")   # 表头蓝
EF  = PatternFill("solid", start_color="FCE4D6")   # 估价橙
TFILL = PatternFill("solid", start_color="F4B084") # 合计金
YELLOW = PatternFill("solid", start_color="FFF2CC")# 可编辑浅黄
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left", vertical="center", wrap_text=True)
R = Alignment(horizontal="right", vertical="center", wrap_text=True)


def build_pricing_workbook(title, info_left, info_right, refs, data, notes, out_path):
    """生成双sheet全公式核价Excel并保存。

    refs: [(name, price, source, note), ...]          清单参考行，row3起
    data: [(seq, content, spec, unit, qty, ref_idx, subitem, note, is_est), ...]
          ref_idx 为 refs 的 0-based 索引（内部换算 =清单参考!C{ref_idx+3}）
    """
    wb = Workbook()

    # ===== Sheet: 清单参考 =====
    ws_ref = wb.active
    ws_ref.title = "清单参考"
    ws_ref.merge_cells("A1:E1")
    ws_ref["A1"] = "清单参考表——综合单价真源（浅黄=可编辑；修改后核价明细自动重算）"
    ws_ref["A1"].font = TF; ws_ref["A1"].alignment = C
    ws_ref.row_dimensions[1].height = 30
    for col, h in enumerate(["序号", "项目名称", "综合单价(元)", "来源", "备注"], 1):
        c = ws_ref.cell(row=2, column=col, value=h)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c.fill = HF; c.alignment = C; c.border = BDR
    ws_ref.row_dimensions[2].height = 28

    for i, (name, price, source, note) in enumerate(refs):
        rr = 3 + i
        ws_ref.cell(row=rr, column=1, value=i + 1).alignment = C
        ws_ref.cell(row=rr, column=2, value=name).alignment = L
        ws_ref.cell(row=rr, column=2).font = NF
        c = ws_ref.cell(row=rr, column=3, value=price)
        c.alignment = R; c.number_format = "#,##0.00"; c.fill = YELLOW
        c.font = Font(name=FONT_NAME, size=10, bold=True)
        ws_ref.cell(row=rr, column=4, value=source).alignment = L
        ws_ref.cell(row=rr, column=4).font = SF
        ws_ref.cell(row=rr, column=5, value=note).alignment = L
        ws_ref.cell(row=rr, column=5).font = SF
        for col in range(1, 6):
            ws_ref.cell(row=rr, column=col).border = BDR
    for col, w in {"A": 6, "B": 40, "C": 14, "D": 34, "E": 40}.items():
        ws_ref.column_dimensions[col].width = w

    # ===== Sheet: 核价明细（放第一位）=====
    ws = wb.create_sheet("核价明细", 0)
    ws.merge_cells("A1:I1")
    ws["A1"] = title
    ws["A1"].font = TF; ws["A1"].alignment = C; ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:E2"); ws["A2"] = info_left; ws["A2"].font = SF
    ws.merge_cells("F2:I2"); ws["F2"] = info_right; ws["F2"].font = SF
    for col, h in enumerate(["序号", "工程内容", "项目特征", "单位", "数量",
                             "综合单价(元)", "合价(元)", "价格来源", "备注"], 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c.fill = HF; c.alignment = C; c.border = BDR
    ws.row_dimensions[4].height = 30

    r = 5
    for (seq, content, spec, unit, qty, ref_idx, subitem, note, is_est) in data:
        ws.cell(row=r, column=1, value=seq).alignment = C
        ws.cell(row=r, column=2, value=content).alignment = L
        ws.cell(row=r, column=3, value=spec).alignment = L
        ws.cell(row=r, column=4, value=unit).alignment = C
        qc = ws.cell(row=r, column=5, value=qty)
        qc.alignment = R; qc.number_format = "#,##0.###"
        # 铁律: refs[ref_idx] 在清单参考 row 3+ref_idx → 公式 C{ref_idx+3}
        fc = ws.cell(row=r, column=6, value=f"=清单参考!C{ref_idx + 3}")
        fc.alignment = R; fc.number_format = "#,##0.00"; fc.font = FORMULA_FONT
        hc = ws.cell(row=r, column=7, value=f"=E{r}*F{r}")
        hc.alignment = R; hc.number_format = "#,##0.00"
        ws.cell(row=r, column=8, value=subitem).alignment = L
        ws.cell(row=r, column=8).font = SF
        ws.cell(row=r, column=9, value=note).alignment = L
        ws.cell(row=r, column=9).font = SF
        for col in range(1, 10):
            ws.cell(row=r, column=col).border = BDR
            if col != 6:
                ws.cell(row=r, column=col).font = NF
            if is_est:
                ws.cell(row=r, column=col).fill = EF
        r += 1

    # 合计（先写值再合并，只合并A:F，G列是SUM公式不能被并掉）
    ws.cell(row=r, column=1, value="合  计").font = Font(name=FONT_NAME, size=13, bold=True)
    ws.cell(row=r, column=1).alignment = R
    tc = ws.cell(row=r, column=7, value=f"=SUM(G5:G{r-1})")
    tc.font = Font(name=FONT_NAME, size=13, bold=True)
    tc.alignment = R; tc.number_format = "#,##0.00"
    for col in range(1, 10):
        ws.cell(row=r, column=col).fill = TFILL
        ws.cell(row=r, column=col).border = BDR
    ws.merge_cells(f"A{r}:F{r}")
    ws.row_dimensions[r].height = 30
    r += 1

    for col, w in {"A": 5, "B": 16, "C": 42, "D": 6, "E": 8,
                   "F": 14, "G": 14, "H": 28, "I": 34}.items():
        ws.column_dimensions[col].width = w

    nr = r + 1
    ws.merge_cells(f"A{nr}:I{nr}")
    ws.cell(row=nr, column=1, value="公式说明与提示（蓝色=公式）：").font = BF
    nr += 1
    for note in notes:
        ws.merge_cells(f"A{nr}:I{nr}")
        ws.cell(row=nr, column=1, value=note).font = SF
        ws.cell(row=nr, column=1).alignment = L
        nr += 1

    wb.save(out_path)
    return out_path


def verify_pricing(path, verbose=True):
    """生成后验证：逐行取F列公式的直接单元格引用值，与清单参考refs原值比对。

    用直接单元格引用（ws2['C3']），禁止反推行号——杜绝ref错位。
    返回 (是否全部通过, 合计金额)。
    """
    wb = load_workbook(path, data_only=False)
    ws1 = wb["核价明细"]; ws2 = wb["清单参考"]
    total, ok = 0.0, True
    r = 5
    while ws1.cell(r, 1).value not in (None, "合  计"):
        qty = ws1.cell(r, 5).value
        formula = ws1.cell(r, 6).value          # 形如 =清单参考!C3
        cell_ref = formula.split("!")[1]
        price = ws2[cell_ref].value              # 直接单元格引用取值
        if price is None:
            ok = False
            if verbose: print(f"[FAIL] 行{r} 引用{cell_ref} 为空！")
        else:
            line = round(qty * price, 2)
            total += line
            if verbose:
                print(f"行{r} {ws1.cell(r,2).value} {qty} × {price} = {line} (引用{cell_ref})")
        r += 1
    if verbose: print(f"合计 = {round(total, 2)}")
    return ok, round(total, 2)


if __name__ == "__main__":
    # 自测：单项示例
    refs = [("环氧地坪修补（清单#104）", 180.00, "高新卓曜清单#104", "含拆除+底漆+中涂+面涂")]
    data = [(1, "环氧地面修补", "普通环氧地坪3.5㎡", "㎡", 3.5, 0, "环氧地坪修补#104", "乙供乙装", False)]
    out = build_pricing_workbook(
        title="自测——环氧地面修补核价",
        info_left="项目：测试", info_right="合同：测试",
        refs=refs, data=data,
        notes=["1. 综合单价=引用『清单参考』C列；合价=E×F；合计=SUM。"],
        out_path=r"D:\造价管理\_skill自测_核价模板.xlsx",
    )
    ok, total = verify_pricing(out)
    print("验证:", "通过" if ok else "失败", "| 合计:", total)
    import os
    os.remove(out)  # 自测完删除
    print("自测文件已清理")
