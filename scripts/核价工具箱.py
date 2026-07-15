"""
零星工程核价表生成器
用法：py make_pricelist.py <审批单路径> [输出路径]
"""
import openpyxl, sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============ 清单单价数据库（可直接引用）============
PRICE_DB = {
    "管道疏通":         {"sub": "卫生间内下水管道疏通",    "unit": "次", "price": 97.48, "src": "土建及修缮"},
    "小便感应器":        {"sub": "小便感应器调换（感应延迟阀，暗装直流，乙供乙装）", "unit": "组", "price": 500.00, "src": "电气配件安装及拆除"},
    "小便延迟阀":        {"sub": "小便延迟阀调换（乙供乙装）", "unit": "个", "price": 134.00, "src": "土建及修缮"},
    "T8灯管":           {"sub": "T8日光灯管更换（36W，乙供乙装）", "unit": "根", "price": 30.00, "src": "灯具安装及拆除"},
    "平板灯":           {"sub": "平板灯安装（吸顶式，乙供乙装）", "unit": "个", "price": 40.00, "src": "灯具安装及拆除"},
    "木门锁":           {"sub": "木门锁拆换（乙供乙装）",   "unit": "套", "price": 64.62, "src": "土建及修缮"},
    "锁芯":            {"sub": "木门锁锁芯拆换（乙供乙装）", "unit": "个", "price": 36.41, "src": "土建及修缮"},
    "合页":            {"sub": "合页拆换（不锈钢，乙供乙装）","unit": "个", "price": 36.41, "src": "土建及修缮"},
    "瓷砖铺贴":          {"sub": "内墙面砖铺贴（人工，技工）","unit": "m²","price": 168.00,"src": "人工报价及费率"},
    "拆除台盆":          {"sub": "拆除台盆、拖把池（乙供乙装）","unit": "套", "price": 34.02, "src": "土建及修缮"},
    "大便延迟阀":        {"sub": "大便延迟阀调换（乙供乙装）","unit": "个","price": 236.00,"src": "土建及修缮"},
}

# 技工/普工工日单价
LABOR = {"技工": 420.00, "普工": 280.00}

def match_item(desc: str) -> dict:
    """根据维修描述匹配清单子目"""
    d = desc.upper()
    # 感应关键词 → 小便感应器
    if "小便" in d and ("感应" in d or "红外" in d):
        return {**PRICE_DB["小便感应器"], "matched": "小便感应器"}
    # 小便+不出水/下水管脱落 → 感应器（常见故障，优先感应）
    if "小便" in d and ("不出水" in d or "下水管" in d or "脱落" in d or "不下水" in d):
        return {**PRICE_DB["小便感应器"], "matched": "小便感应器(故障推断)"}
    # 普通小便阀
    if "小便" in d and "阀" in d:
        return {**PRICE_DB["小便延迟阀"], "matched": "小便延迟阀"}
    if "大便" in d and "阀" in d:
        return {**PRICE_DB["大便延迟阀"], "matched": "大便延迟阀"}
    if "疏通" in d or "堵塞" in d:
        return {**PRICE_DB["管道疏通"], "matched": "管道疏通"}
    if "灯管" in d:
        return {**PRICE_DB["T8灯管"], "matched": "T8灯管"}
    if "吸顶灯" in d or "灯壳" in d or "平板灯" in d:
        return {**PRICE_DB["平板灯"], "matched": "平板灯"}
    if "门锁" in d and "锁芯" not in d and "芯" not in d:
        return {**PRICE_DB["木门锁"], "matched": "木门锁"}
    if "锁芯" in d or "芯坏" in d or "芯坏" in d:
        return {**PRICE_DB["锁芯"], "matched": "锁芯"}
    if "合页" in d or "折页" in d or "铰链" in d:
        return {**PRICE_DB["合页"], "matched": "合页"}
    if "瓷砖" in d or "墙砖" in d or "砖脱落" in d or "面砖" in d:
        return {**PRICE_DB["瓷砖铺贴"], "matched": "瓷砖铺贴"}
    if "拆除" in d and ("水池" in d or "台盆" in d or "拖把池" in d):
        return {**PRICE_DB["拆除台盆"], "matched": "拆除台盆"}
    # 无匹配 → 需人工判断
    return None

def estimate_price(desc: str, qty: float) -> dict:
    """无法匹配时给出估价建议"""
    return {"matched": "⚠️ 需人工估价", "price": 0, "unit": "项", "src": "待确认",
            "tip": "建议参考技工420元/工日、普工280元/工日估算"}

# ============ Excel 样式 ============
def make_xlsx(items: list, src_name: str, out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "核价明细"

    hf = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    hb = PatternFill("solid", fgColor="4472C4")
    tf = Font(name="Arial", bold=True, size=14)
    sf = Font(name="Arial", size=10, italic=True)
    nf = Font(name="Arial", size=10)
    totf = Font(name="Arial", bold=True, size=11)
    c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    bd = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
    wf = PatternFill("solid", fgColor="FFF2CC")
    tf2 = PatternFill("solid", fgColor="D9E1F2")

    ws.merge_cells("A1:I1")
    ws["A1"] = "零星工程维修核价明细表"
    ws["A1"].font = tf; ws["A1"].alignment = c
    ws.merge_cells("A2:I2")
    ws["A2"] = f"依据：徐州光伏2026年度零星工程施工中标清单v2  |  来源：{src_name}  |  合同：单价合同（下浮10%已含，结算不乘0.9）"
    ws["A2"].font = sf; ws["A2"].alignment = c

    headers = ["序号","报修业务","故障区域","维修内容（原文）","清单子目","单位","工程量","综合单价(元)","合价(元)"]
    for col, h in enumerate(headers, 1):
        c_ = ws.cell(row=4, column=col, value=h)
        c_.font = hf; c_.fill = hb; c_.alignment = c; c_.border = bd

    for i, item in enumerate(items):
        row = 5 + i
        # item: (seq, biz, area, desc, match_result)
        seq, biz, area, desc, mr = item
        sub  = mr.get("sub", mr.get("matched",""))
        unit = mr.get("unit", "项")
        qty2 = float(item[6] if len(item) > 6 else 1)
        price = mr.get("price", 0)
        matched = mr.get("matched","")
        remark = mr.get("src","")
        is_warn = "⚠️" in matched or "需人工" in matched

        row_data = [seq, biz, area, desc, sub, unit, qty2, price, f"=G{row}*H{row}"]
        for col, val in enumerate(row_data, 1):
            c_ = ws.cell(row=row, column=col, value=val)
            c_.font = nf; c_.border = bd
            if col in (1,6,7,8,9):
                c_.alignment = c; c_.number_format = "0.00"
            else:
                c_.alignment = l
            if is_warn:
                c_.fill = wf

    tot_row = 5 + len(items)
    for col in range(1,10):
        c_ = ws.cell(row=tot_row, column=col)
        c_.border = bd; c_.fill = tf2
        if col == 1:
            c_.value = "合计"; c_.font = totf; c_.alignment = c
        elif col == 9:
            c_.value = f"=SUM(I5:I{tot_row-1})"; c_.font = totf
            c_.alignment = c; c_.number_format = "0.00"
        elif col == 8:
            c_.value = ""; c_.font = totf; c_.alignment = c
    ws.merge_cells(f"A{tot_row}:G{tot_row}")

    # 含税
    tx_row = tot_row + 1
    ws.merge_cells(f"A{tx_row}:G{tx_row}")
    for col in range(1,10):
        c_ = ws.cell(row=tx_row, column=col)
        c_.border = bd; c_.fill = tf2
        if col == 1:
            c_.value = "含税合计（9%）"; c_.font = totf; c_.alignment = c
        elif col == 9:
            c_.value = f"=I{tot_row}*1.09"; c_.font = totf
            c_.alignment = c; c_.number_format = "0.00"
    ws.cell(row=tx_row, column=8).value = ""
    ws.cell(row=tx_row, column=1).border = bd

    # 列宽
    for i, w in enumerate([6,10,14,46,38,8,10,14,14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[4].height = 28

    # 说明
    nr = tx_row + 2
    ws.merge_cells(f"A{nr}:I{nr}")
    ws.cell(row=nr, column=1, value="说明：").font = Font(name="Arial", bold=True, size=10)
    ws.merge_cells(f"A{nr+1}:I{nr+1}")
    ws.cell(row=nr+1, column=1, value="⚠️ 标记项为清单中无精确匹配子目，需人工估价，实际价格以现场核验为准。").font = Font(name="Arial", size=9)
    ws.merge_cells(f"A{nr+2}:I{nr+2}")
    ws.cell(row=nr+2, column=1, value="合同规则：单价合同，下浮10%已含在中标单价中，结算不再乘0.9。").font = Font(name="Arial", size=9)

    wb.save(out_path)
    print(f"✅ 已保存: {out_path}")
    total = sum((item[6] if len(item)>6 else 1) * item[4].get("price",0) for item in items)
    print(f"合计: {total:.2f} 元 | 含税: {total*1.09:.2f} 元")


def read_approval(path: str) -> list:
    """读取审批单，尝试提取维修项"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        items = []
        for row in range(4, ws.max_row + 1):
            vals = [ws.cell(row=row, column=c).value for c in range(1, 10)]
            if any(vals):
                items.append(vals)
        return items
    except Exception as e:
        print(f"读取失败: {e}")
        return []

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: py make_pricelist.py <审批单路径> [输出路径]")
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    src_name = os.path.basename(src)
    if out is None:
        fname = f"核价明细_{src_name.replace('.xlsx','')}_v2.xlsx"
        out = os.path.join(r"C:\Users\山川\Desktop\龙虾同步文件夹", fname)

    items = read_approval(src)
    print(f"读取到 {len(items)} 行数据")
    # TODO: 接入山川大人的中标清单进行匹配
    print("（请接入中标清单v2进行子目匹配）")
