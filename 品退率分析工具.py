#!/usr/bin/env python3
"""
品退率横向对比分析工具
读取各店铺品退率Excel文件，生成横向对比汇总表
适用平台：Windows / macOS
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import json
import threading
from pathlib import Path

# ====================== 配置文件 ======================
CONFIG_FILE = Path(__file__).parent / "品退率分析工具_config.json"

DEFAULT_STORE_KEYWORDS = {
    "鲜采水果专营店": ["鲜采水果"],
    "金榜水果专营店": ["金榜水果"],
    "山姆伯伯生鲜旗舰店": ["山姆伯伯生鲜"],
    "爱沃尔德果蔬专营店": ["爱沃尔德果蔬专营店", "爱沃尔德"],
    "金枕榴莲产地直达店": ["金枕榴莲产地直达", "金枕榴莲"],
    "山姆大叔果蔬旗舰店": ["山姆大叔果蔬旗舰店", "山姆大叔果蔬"],
    "山姆大叔水果旗舰店": ["山姆大叔水果"],
    "山姆大叔生鲜旗舰店": ["山姆大叔生鲜"],
}

DEFAULT_PERIOD_KEYWORDS = {
    "7天": ["7天", "近7天"],
    "30天": ["30天", "近30天"],
}


def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ====================== 核心逻辑 ======================
def safe_int(v):
    if v is None or v == "" or v == "-":
        return 0
    try:
        return int(v)
    except (ValueError, TypeError):
        return 0


def safe_float(v):
    if v is None or v == "" or v == "-":
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def match_store(filename, store_keywords):
    """根据文件名中的关键词匹配店铺"""
    for store_name, keywords in store_keywords.items():
        for kw in keywords:
            if kw in filename:
                return store_name
    return None


def match_period(filename, period_keywords):
    """根据文件名匹配时间周期"""
    for period, keywords in period_keywords.items():
        for kw in keywords:
            if kw in filename:
                return period
    return None


def scan_folder(folder_path, store_keywords, period_keywords):
    """扫描文件夹，自动匹配文件到店铺和周期"""
    matches = {}  # {store: {period: filepath}}
    errors = []

    for fname in os.listdir(folder_path):
        if not fname.endswith(".xlsx"):
            continue
        filepath = os.path.join(folder_path, fname)

        store = match_store(fname, store_keywords)
        period = match_period(fname, period_keywords)

        if store is None:
            errors.append(f"无法识别店铺: {fname}")
            continue
        if period is None:
            errors.append(f"无法识别周期(7天/30天): {fname}")
            continue

        if store not in matches:
            matches[store] = {}
        matches[store][period] = filepath

    return matches, errors


def load_all_data(matches):
    """加载所有文件数据"""
    data = {}
    for store, periods in matches.items():
        data[store] = {}
        for period, filepath in periods.items():
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
            period_data = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                cat1, cat2, cat3, qty, rate, weight, industry = row
                if cat3:
                    period_data[str(cat3)] = {
                        "qty": safe_int(qty),
                        "rate": safe_float(rate),
                        "weight": safe_float(weight) or 0,
                        "industry": str(industry) if industry else "",
                    }
            data[store][period] = period_data
    return data


def calc_internal_avg(data, stores, cat, period):
    total_qty = 0
    weighted_sum = 0
    for store in stores:
        d = data[store][period].get(cat, {})
        qty = d.get("qty", 0)
        rate = d.get("rate")
        if qty > 0 and rate is not None:
            total_qty += qty
            weighted_sum += rate * qty
    if total_qty > 0:
        return weighted_sum / total_qty, total_qty
    return None, 0


def has_sales(d):
    return d.get("qty", 0) > 0 and d.get("rate") is not None


def generate_excel(data, stores, output_path, progress_callback=None):
    """生成汇总Excel"""
    # Collect categories
    all_cats_7 = set()
    all_cats_30 = set()
    for store in stores:
        if "7天" in data[store]:
            all_cats_7.update(data[store]["7天"].keys())
        if "30天" in data[store]:
            all_cats_30.update(data[store]["30天"].keys())

    def sort_cats(cats, period):
        cat_vol = {}
        for c in cats:
            _, total = calc_internal_avg(data, stores, c, period)
            cat_vol[c] = total
        return sorted(cats, key=lambda x: cat_vol[x], reverse=True)

    sorted_cats_7 = sort_cats(all_cats_7, "7天")
    sorted_cats_30 = sort_cats(all_cats_30, "30天")

    # Styles
    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    sub_header_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    sub_header_font = Font(name="微软雅黑", bold=True, size=9)
    cat_font = Font(name="微软雅黑", bold=True, size=10)
    data_font = Font(name="微软雅黑", size=10)
    avg_font = Font(name="微软雅黑", bold=True, size=10, color="2F5496")
    na_font = Font(name="微软雅黑", size=9, color="AAAAAA")

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    orange_fill = PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid")
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    avg_fill = PatternFill(start_color="E8EEF7", end_color="E8EEF7", fill_type="solid")
    na_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def industry_fill(level, has_sales_flag):
        if not has_sales_flag:
            return na_fill
        s = str(level)
        if "优于" in s: return green_fill
        if "1-2倍" in s: return yellow_fill
        if "2-3倍" in s: return orange_fill
        if "3-5倍" in s or "5倍以上" in s: return red_fill
        return white_fill

    def industry_label(level, has_sales_flag):
        if not has_sales_flag:
            return "无销售"
        s = str(level)
        if "优于" in s: return "达标"
        if "1-2倍" in s: return "轻微超标"
        if "2-3倍" in s: return "超标"
        if "3-5倍" in s or "5倍以上" in s: return "严重超标"
        return "-"

    def label_color(label):
        if label == "达标": return "008000"
        if label == "轻微超标": return "806000"
        if label == "超标": return "C00000"
        if label == "严重超标": return "C00000"
        if label == "无销售": return "AAAAAA"
        return "000000"

    if progress_callback:
        progress_callback("正在创建近7天汇总...")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---- Sheet 1 & 2 ----
    for sheet_name, period, sorted_cats in [
        ("近7天品退率汇总", "7天", sorted_cats_7),
        ("近30天品退率汇总", "30天", sorted_cats_30),
    ]:
        ws = wb.create_sheet(title=sheet_name)
        n_stores = len(stores)
        total_cols = 2 + n_stores * 2

        # Title
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        c = ws.cell(row=1, column=1, value=f"各店铺品类品退率横向对比 — {sheet_name}")
        c.font = Font(name="微软雅黑", bold=True, size=14, color="2F5496")
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32
        ws.row_dimensions[2].height = 6

        # Row 3: Header
        for c_idx in range(1, total_cols + 1):
            cell = ws.cell(row=3, column=c_idx)
            cell.fill = header_fill
            cell.border = thin_border
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=3, column=1, value="三级品类")
        ws.cell(row=3, column=2, value="内部加权均值")

        col = 3
        for store in stores:
            ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
            ws.cell(row=3, column=col, value=store)
            col += 2
        ws.row_dimensions[3].height = 24

        # Row 4: Sub-header
        for c_idx in range(1, total_cols + 1):
            cell = ws.cell(row=4, column=c_idx)
            if c_idx <= 2:
                cell.fill = header_fill
            else:
                cell.fill = sub_header_fill
                cell.font = sub_header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        col = 3
        for _ in stores:
            ws.cell(row=4, column=col, value="品退率")
            ws.cell(row=4, column=col + 1, value="达标判定")
            col += 2
        ws.row_dimensions[4].height = 20

        # Data
        row = 5
        for cat in sorted_cats:
            ws.cell(row=row, column=1, value=cat).font = cat_font
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=row, column=1).border = thin_border

            internal_avg, _ = calc_internal_avg(data, stores, cat, period)
            avg_cell = ws.cell(row=row, column=2, value=internal_avg if internal_avg is not None else "-")
            avg_cell.font = avg_font
            avg_cell.alignment = Alignment(horizontal="center", vertical="center")
            avg_cell.border = thin_border
            avg_cell.fill = avg_fill
            if internal_avg is not None:
                avg_cell.number_format = "0.00%"

            col = 3
            for store in stores:
                sd = data[store].get(period, {}).get(cat, {})
                rate = sd.get("rate")
                qty = sd.get("qty", 0)
                industry = sd.get("industry", "")
                hs = has_sales(sd)

                if hs:
                    rc = ws.cell(row=row, column=col, value=rate)
                    rc.number_format = "0.00%"
                else:
                    rc = ws.cell(row=row, column=col, value="-")
                rc.font = na_font if not hs else data_font
                rc.alignment = Alignment(horizontal="center", vertical="center")
                rc.border = thin_border
                rc.fill = industry_fill(industry, hs)

                label = industry_label(industry, hs)
                db = ws.cell(row=row, column=col + 1, value=label)
                db.font = Font(name="微软雅黑", size=9, color=label_color(label))
                db.alignment = Alignment(horizontal="center", vertical="center")
                db.border = thin_border
                db.fill = industry_fill(industry, hs)

                col += 2

            ws.row_dimensions[row].height = 20
            row += 1

        ws.freeze_panes = "C5"
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 14
        for i in range(n_stores):
            ws.column_dimensions[get_column_letter(3 + i * 2)].width = 10
            ws.column_dimensions[get_column_letter(4 + i * 2)].width = 9

    if progress_callback:
        progress_callback("正在创建汇总概览...")

    # ---- Sheet 3: 汇总概览 ----
    ws = wb.create_sheet(title="汇总概览")

    ws.merge_cells("A1:H1")
    ws.cell(row=1, column=1, value="品退率横向对比 — 汇总概览").font = Font(name="微软雅黑", bold=True, size=14, color="2F5496")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A3:F3")
    ws.cell(row=3, column=1, value="颜色图例（按行业水平判定，仅统计有销售的品类）").font = Font(name="微软雅黑", bold=True, size=11)

    legend = [
        (4, "达标 — 优于行业均值", green_fill),
        (5, "轻微超标 — 行业均值 1-2 倍", yellow_fill),
        (6, "超标 — 行业均值 2-3 倍", orange_fill),
        (7, "严重超标 — 行业均值 3-5 倍 / 5倍以上", red_fill),
        (8, "无销售 — 该店铺未经营此品类", na_fill),
    ]
    for r, label, fill in legend:
        c = ws.cell(row=r, column=1, value=f"  {label}")
        c.fill = fill
        c.font = data_font
        c.border = thin_border

    # 严重超标清单
    def write_problem_list(ws, start_row, period, title):
        r = start_row
        ws.merge_cells(f"A{r}:G{r}")
        ws.cell(row=r, column=1, value=title).font = Font(name="微软雅黑", bold=True, size=12, color="C00000")
        r += 1

        headers = ["店铺", "三级品类", "品退率", "行业判定", "订单量", "占店铺订单比", "内部均值参考"]
        for i, h in enumerate(headers):
            c = ws.cell(row=r, column=i + 1, value=h)
            c.font = Font(name="微软雅黑", bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border
        r += 1

        for store in stores:
            if period not in data[store]:
                continue
            sd = data[store][period]
            store_qty = sum(d["qty"] for d in sd.values())
            problem = []
            for cat, d in sd.items():
                if not has_sales(d):
                    continue
                ind = str(d.get("industry", ""))
                if "3-5倍" in ind or "5倍以上" in ind:
                    problem.append((cat, d["rate"], ind, d["qty"]))
            problem.sort(key=lambda x: x[3], reverse=True)

            first = True
            if problem:
                for cat, rate, ind, qty in problem:
                    ws.cell(row=r, column=1, value=store if first else "").font = data_font
                    ws.cell(row=r, column=1).border = thin_border
                    ws.cell(row=r, column=2, value=cat).font = data_font
                    ws.cell(row=r, column=2).border = thin_border
                    rc = ws.cell(row=r, column=3, value=round(rate, 4) if rate else "-")
                    rc.font = data_font; rc.border = thin_border; rc.fill = red_fill
                    rc.number_format = "0.00%"; rc.alignment = Alignment(horizontal="center")
                    ws.cell(row=r, column=4, value=ind).font = data_font
                    ws.cell(row=r, column=4).border = thin_border
                    ws.cell(row=r, column=4).fill = red_fill
                    ws.cell(row=r, column=5, value=qty).font = data_font
                    ws.cell(row=r, column=5).border = thin_border
                    ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
                    pct = f"{qty / store_qty * 100:.1f}%" if store_qty > 0 else "-"
                    ws.cell(row=r, column=6, value=pct).font = data_font
                    ws.cell(row=r, column=6).border = thin_border
                    ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")
                    int_avg, _ = calc_internal_avg(data, stores, cat, period)
                    ref = f"{int_avg:.2%}" if int_avg else "-"
                    ws.cell(row=r, column=7, value=ref).font = Font(name="微软雅黑", size=9, color="2F5496")
                    ws.cell(row=r, column=7).border = thin_border
                    ws.cell(row=r, column=7).alignment = Alignment(horizontal="center")
                    first = False
                    r += 1
            else:
                ws.cell(row=r, column=1, value=store).font = data_font
                ws.cell(row=r, column=1).border = thin_border
                gc = ws.cell(row=r, column=2, value="无严重超标品类")
                gc.font = Font(name="微软雅黑", size=10, color="008000")
                gc.fill = green_fill
                ws.cell(row=r, column=2).border = thin_border
                r += 1
        return r + 1

    r = 10
    r = write_problem_list(ws, r, "30天", "严重超标品类清单（行业均值3-5倍及以上）— 近30天")
    r = write_problem_list(ws, r, "7天", "严重超标品类清单（行业均值3-5倍及以上）— 近7天")

    # 店铺健康度
    def write_health(ws, start_row, period, title):
        r = start_row
        ws.merge_cells(f"A{r}:H{r}")
        ws.cell(row=r, column=1, value=title).font = Font(name="微软雅黑", bold=True, size=12, color="2F5496")
        r += 1

        headers = ["店铺", "经营品类数", "达标", "轻微超标", "超标", "严重超标", "总订单量", "店铺加权品退率"]
        for i, h in enumerate(headers):
            c = ws.cell(row=r, column=i + 1, value=h)
            c.font = Font(name="微软雅黑", bold=True, size=9, color="FFFFFF")
            c.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border
        r += 1

        for store in stores:
            if period not in data[store]:
                continue
            sd = data[store][period]
            active = {c: d for c, d in sd.items() if has_sales(d)}
            total = len(active)
            dabiao = sum(1 for d in active.values() if "优于" in str(d.get("industry", "")))
            qingwei = sum(1 for d in active.values() if "1-2倍" in str(d.get("industry", "")))
            chaobiao = sum(1 for d in active.values() if "2-3倍" in str(d.get("industry", "")))
            yanzhong = sum(1 for d in active.values() if "3-5倍" in str(d.get("industry", "")) or "5倍以上" in str(d.get("industry", "")))
            total_qty = sum(d.get("qty", 0) for d in active.values())

            w_sum = 0; w_qty = 0
            for d in active.values():
                if d["rate"] is not None and d["qty"] > 0:
                    w_sum += d["rate"] * d["qty"]
                    w_qty += d["qty"]
            store_wavg = w_sum / w_qty if w_qty > 0 else None

            vals = [store, total, dabiao, qingwei, chaobiao, yanzhong, total_qty, store_wavg]
            for i, val in enumerate(vals):
                cell = ws.cell(row=r, column=i + 1, value=val)
                cell.font = data_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
                if i == 7 and val is not None:
                    cell.number_format = "0.00%"

            ws.cell(row=r, column=3).fill = green_fill
            if qingwei > 0: ws.cell(row=r, column=4).fill = yellow_fill
            if chaobiao > 0: ws.cell(row=r, column=5).fill = orange_fill
            if yanzhong > 0: ws.cell(row=r, column=6).fill = red_fill

            r += 1
        return r + 1

    r = write_health(ws, r, "30天", "店铺健康度总览 — 近30天（仅统计有销售的品类）")
    r = write_health(ws, r, "7天", "店铺健康度总览 — 近7天（仅统计有销售的品类）")

    for col, w in [("A", 22), ("B", 16), ("C", 14), ("D", 16), ("E", 13), ("F", 16), ("G", 12), ("H", 16)]:
        ws.column_dimensions[col].width = w

    if progress_callback:
        progress_callback("正在保存文件...")

    wb.save(output_path)


# ====================== GUI ======================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("品退率横向对比分析工具 v1.0")
        self.root.geometry("800x680")
        self.root.minsize(700, 600)

        # 设置中文字体
        self.default_font = ("微软雅黑", 10)
        self.title_font = ("微软雅黑", 14, "bold")
        self.root.option_add("*Font", self.default_font)

        # 样式
        style = ttk.Style()
        style.theme_use("clam")

        # 变量
        self.input_folder = tk.StringVar(value="")
        self.output_folder = tk.StringVar(value="")
        self.store_keywords = DEFAULT_STORE_KEYWORDS.copy()
        self.period_keywords = DEFAULT_PERIOD_KEYWORDS.copy()

        # 加载配置
        cfg = load_config()
        if "input_folder" in cfg:
            self.input_folder.set(cfg["input_folder"])
        if "output_folder" in cfg:
            self.output_folder.set(cfg["output_folder"])
        if "store_keywords" in cfg:
            self.store_keywords = cfg["store_keywords"]
        if "period_keywords" in cfg:
            self.period_keywords = cfg["period_keywords"]

        self.build_ui()

    def build_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(main_frame, text="品退率横向对比分析工具", font=self.title_font)
        title_label.pack(pady=(0, 15))

        # ---- 文件夹选择 ----
        folder_frame = ttk.LabelFrame(main_frame, text="文件路径设置", padding=10)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        # 输入文件夹
        ttk.Label(folder_frame, text="源文件目录:").grid(row=0, column=0, sticky=tk.W, pady=5)
        input_entry = ttk.Entry(folder_frame, textvariable=self.input_folder, width=60)
        input_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(folder_frame, text="选择...", command=self.select_input).grid(row=0, column=2, pady=5)

        # 输出目录
        ttk.Label(folder_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(folder_frame, textvariable=self.output_folder, width=60)
        output_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(folder_frame, text="选择...", command=self.select_output).grid(row=1, column=2, pady=5)

        folder_frame.columnconfigure(1, weight=1)

        # ---- 操作按钮 ----
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.scan_btn = ttk.Button(btn_frame, text="1. 扫描文件", command=self.scan_files)
        self.scan_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.generate_btn = ttk.Button(btn_frame, text="2. 生成汇总表", command=self.generate, state=tk.DISABLED)
        self.generate_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 进度条
        self.progress = ttk.Progressbar(btn_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        # ---- 文件匹配结果 ----
        result_frame = ttk.LabelFrame(main_frame, text="文件匹配结果", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Treeview
        columns = ("店铺", "7天文件", "30天文件", "状态")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=8)
        for col in columns:
            self.tree.heading(col, text=col)
        self.tree.column("店铺", width=180)
        self.tree.column("7天文件", width=250)
        self.tree.column("30天文件", width=250)
        self.tree.column("状态", width=80)

        tree_scroll = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # ---- 日志 ----
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=6, wrap=tk.WORD, font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # ---- 底部 ----
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(bottom_frame, text="配置店铺关键词", command=self.open_config).pack(side=tk.LEFT)

        self.status_label = ttk.Label(bottom_frame, text="就绪", foreground="gray")
        self.status_label.pack(side=tk.RIGHT)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def set_status(self, msg, color="gray"):
        self.status_label.config(text=msg, foreground=color)

    def select_input(self):
        path = filedialog.askdirectory(title="选择包含各店铺品退率Excel的文件夹")
        if path:
            self.input_folder.set(path)
            self.save_current_config()

    def select_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_folder.set(path)
            self.save_current_config()

    def save_current_config(self):
        cfg = {
            "input_folder": self.input_folder.get(),
            "output_folder": self.output_folder.get(),
            "store_keywords": self.store_keywords,
            "period_keywords": self.period_keywords,
        }
        save_config(cfg)

    def scan_files(self):
        input_path = self.input_folder.get()
        if not input_path or not os.path.isdir(input_path):
            messagebox.showerror("错误", "请先选择有效的源文件目录")
            return

        self.progress.start()
        self.set_status("正在扫描...")
        self.log("========== 开始扫描 ==========")

        self.matches, errors = scan_folder(input_path, self.store_keywords, self.period_keywords)

        # 更新tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        all_stores = set(list(self.store_keywords.keys()) + list(self.matches.keys()))
        for store in sorted(all_stores):
            files_7 = self.matches.get(store, {}).get("7天", "")
            files_30 = self.matches.get(store, {}).get("30天", "")
            if files_7:
                files_7 = os.path.basename(files_7)
            if files_30:
                files_30 = os.path.basename(files_30)

            status = "完整" if files_7 and files_30 else ("缺7天" if not files_7 else "缺30天")
            self.tree.insert("", tk.END, values=(store, files_7 or "未找到", files_30 or "未找到", status))

        for err in errors:
            self.log(f"[警告] {err}")

        complete = sum(1 for store in self.matches if len(self.matches[store]) == 2)
        total = len(self.matches)
        self.log(f"扫描完成: {complete}/{total} 家店铺数据完整")

        if complete >= 1:
            self.generate_btn.config(state=tk.NORMAL)
            self.set_status(f"扫描完成，{complete}家可生成", "green")
        else:
            self.generate_btn.config(state=tk.DISABLED)
            self.set_status("扫描完成，无完整数据", "red")

        self.progress.stop()
        self.save_current_config()

    def generate(self):
        input_path = self.input_folder.get()
        output_path = self.output_folder.get()

        if not input_path or not os.path.isdir(input_path):
            messagebox.showerror("错误", "请先选择源文件目录")
            return
        if not output_path or not os.path.isdir(output_path):
            messagebox.showerror("错误", "请先选择输出目录")
            return

        # Re-scan
        self.matches, errors = scan_folder(input_path, self.store_keywords, self.period_keywords)
        stores = sorted(self.matches.keys())
        # Only include stores with at least one period
        stores = [s for s in stores if len(self.matches[s]) >= 1]

        if not stores:
            messagebox.showerror("错误", "没有找到任何有效数据")
            return

        output_file = os.path.join(output_path, "品退率横向对比汇总.xlsx")

        self.progress.start()
        self.generate_btn.config(state=tk.DISABLED)

        def run():
            try:
                data = load_all_data(self.matches)

                def progress_cb(msg):
                    self.log(f"[进度] {msg}")
                    self.set_status(msg)

                self.log("========== 开始生成 ==========")
                self.log(f"涉及店铺: {', '.join(stores)}")

                generate_excel(data, stores, output_file, progress_callback=progress_cb)

                self.log(f"✅ 生成完成: {output_file}")
                self.set_status("生成完成!", "green")
                self.root.after(0, lambda: messagebox.showinfo("完成", f"汇总表已生成:\n{output_file}"))
            except Exception as e:
                self.log(f"[错误] {str(e)}")
                self.set_status("生成失败", "red")
                self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
            finally:
                self.progress.stop()
                self.generate_btn.config(state=tk.NORMAL)

        threading.Thread(target=run, daemon=True).start()

    def open_config(self):
        config_win = tk.Toplevel(self.root)
        config_win.title("店铺关键词配置")
        config_win.geometry("600x500")
        config_win.minsize(500, 400)

        frame = ttk.Frame(config_win, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="配置店铺名称与文件名关键词的匹配关系", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(0, 10))
        ttk.Label(frame, text="每个店铺可配置多个关键词（逗号分隔），扫描时会按关键词匹配文件", foreground="gray").pack(anchor=tk.W, pady=(0, 10))

        # 表格
        columns = ("店铺名称", "匹配关键词")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=12)
        tree.heading("店铺名称", text="店铺名称")
        tree.heading("匹配关键词", text="匹配关键词（逗号分隔）")
        tree.column("店铺名称", width=200)
        tree.column("匹配关键词", width=350)

        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        for store, keywords in self.store_keywords.items():
            tree.insert("", tk.END, values=(store, ", ".join(keywords)))

        def save_changes():
            new_config = {}
            for item in tree.get_children():
                values = tree.item(item, "values")
                store_name = values[0].strip()
                keywords = [kw.strip() for kw in values[1].split(",") if kw.strip()]
                if store_name and keywords:
                    new_config[store_name] = keywords
            if new_config:
                self.store_keywords = new_config
                self.save_current_config()
                self.log("店铺关键词配置已更新")
                messagebox.showinfo("已保存", "配置已保存，下次扫描生效")
                config_win.destroy()

        def add_row():
            tree.insert("", tk.END, values=("新店铺", "关键词1, 关键词2"))

        def delete_row():
            sel = tree.selection()
            for item in sel:
                tree.delete(item)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btn_frame, text="+ 添加店铺", command=add_row).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="- 删除选中", command=delete_row).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="保存并关闭", command=save_changes).pack(side=tk.RIGHT)


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
