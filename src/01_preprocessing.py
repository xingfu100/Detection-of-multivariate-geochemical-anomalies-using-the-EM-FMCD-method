"""Geochemical data preprocessing tool.

The first two columns are coordinates and subsequent columns contain element
values. Supports log10, CLR, ILR and standardization using a parameter dialog.
Missing, zero and negative values are replaced by half the detection limit when
available, or half the column minimum positive value otherwise. If several
transforms are selected, priority is ILR > CLR > log10. Outputs include the
complete results, modeling-ready data and a preprocessing report."""

from __future__ import annotations

import os
import re
import sys
import subprocess
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


# ============================================================
# 1. General utilities
# ============================================================

def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def safe_name(name: str) -> str:
    name = str(name).strip()
    if name == "":
        name = "未命名结果"
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", "_", name)
    return name[:120]


def safe_sheet_name(name: str) -> str:
    name = str(name)
    for ch in ["\\", "/", "?", "*", "[", "]", ":"]:
        name = name.replace(ch, "_")
    return name[:31]


def make_unique_columns(columns) -> list[str]:
    seen = {}
    new_cols = []

    for col in columns:
        col_str = str(col).strip()
        if col_str == "" or col_str.lower() == "nan":
            col_str = "未命名列"

        if col_str not in seen:
            seen[col_str] = 0
            new_cols.append(col_str)
        else:
            seen[col_str] += 1
            new_cols.append(f"{col_str}_重复{seen[col_str]}")

    return new_cols


def read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        df = pd.read_excel(path, engine="openpyxl")
    elif suffix == ".xls":
        df = pd.read_excel(path, engine="xlrd")
    elif suffix == ".csv":
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="gbk")
    else:
        raise ValueError(f"暂不支持的文件类型：{suffix}")

    df = df.copy()
    df.columns = make_unique_columns(df.columns)
    return df


def split_geochem_table(df: pd.DataFrame):
    if df.shape[1] < 3:
        raise ValueError("数据至少需要3列：前两列为坐标，第三列以后为元素。")

    coord_cols = df.columns[:2].tolist()
    element_cols = df.columns[2:].tolist()

    coords = df[coord_cols].copy()
    elements = df[element_cols].copy()

    return coords, elements, coord_cols, element_cols


def parse_columns(text: str | None, available_cols: list[str], default_all: bool = True) -> list[str]:
    if text is None:
        raise SystemExit("用户取消元素选择，程序结束。")

    text = str(text).strip()

    if text == "":
        if default_all:
            return available_cols
        raise ValueError("未输入元素名。")

    text = text.replace("，", ",").replace("；", ",").replace(";", ",")
    selected = [x.strip() for x in text.split(",") if x.strip()]

    col_map = {str(c).strip().lower(): c for c in available_cols}

    result = []
    missing = []

    for item in selected:
        key = item.lower()
        if key in col_map:
            result.append(col_map[key])
        else:
            missing.append(item)

    if missing:
        raise ValueError(
            "以下元素列在数据中未找到：\n"
            + "、".join(missing)
            + "\n\n请检查列名是否完全一致。"
        )

    return result


def open_output_folder(path: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


# ============================================================
# 2. Stream-sediment detection limits and column-name matching
# ============================================================
#
# Detection limits below were transcribed from the supplied detection-limit table.
# Values retain the scientific notation used in that document:
#   Pb = 2×10^-6
#   SiO2 = 0.1×10^-2
#   Au = 0.3×10^-9
# etc.
#
# Important: no automatic conversion between ppm, ppb, wt.% and mass fraction is performed.
# Input units must match the detection-limit units.
#
# Elements without a detection limit in the source table (such as B, F and P),
# and any variables not matched in the dictionary, fall back to
# one-half of the smallest positive value in the column
# for replacement.
# ============================================================

DETECTION_LIMITS = {
    # Left side of the source table
    "PB": 2e-6,
    "ZN": 10e-6,
    "TH": 4e-6,
    "CU": 1e-6,
    "NI": 2e-6,
    "LA": 30e-6,
    "CO": 1e-6,
    "MN": 30e-6,
    "BA": 50e-6,
    "CR": 15e-6,
    "NB": 5e-6,
    "SR": 5e-6,
    "TI": 100e-6,
    "V": 20e-6,
    "Y": 5e-6,
    "ZR": 10e-6,
    "AL2O3": 0.05e-2,
    "CAO": 0.05e-2,

    # Right side of the source table
    "FE2O3": 0.05e-2,
    "K2O": 0.05e-2,
    "MGO": 0.05e-2,
    "NA2O": 0.05e-2,
    "SIO2": 0.1e-2,
    "AS": 1e-6,
    "SB": 0.2e-6,
    "BI": 0.1e-6,
    "HG": 0.005e-6,
    "AG": 0.02e-6,
    "BE": 0.5e-6,
    "SN": 1e-6,
    "AU": 0.3e-9,
    "LI": 5e-6,
    "CD": 0.1e-6,
    "MO": 0.5e-6,
    "W": 0.5e-6,
    "U": 0.5e-6,
}

# Normalize common subscripts to recognize oxide names such as SiO2 and Al2O3.
_SUBSCRIPT_TRANS = str.maketrans({
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
})


def normalize_element_name(name: str) -> str:
    """Normalize an element or oxide column name for matching DETECTION_LIMITS.
    
    For example Pb(ppm) becomes PB, and SiO2 and SiO₂ both become SIO2."""
    text = str(name).strip().translate(_SUBSCRIPT_TRANS).upper()

    # Remove common units and separators for matching without changing original column names.
    for token in ["PPM", "PPB", "WT.%", "WT%", "MASS%", "%"]:
        text = text.replace(token, "")

    text = re.sub(r"[\\s_\-（）()\[\]{}]", "", text)

    aliases = {
        "TFE2O3": "FE2O3",
        "TOTALFE2O3": "FE2O3",
        "FE2O3T": "FE2O3",
        "TFEO": "FE2O3",
    }
    return aliases.get(text, text)


def build_detection_limit_table(columns: list[str]) -> pd.DataFrame:
    """Build a report table of detection limits and replacement rules for selected variables."""
    rows = []
    for col in columns:
        key = normalize_element_name(col)
        dl = DETECTION_LIMITS.get(key, np.nan)
        if pd.notna(dl):
            method = "1/2检出限"
            default_replacement = float(dl) / 2.0
            source = "用户提供Word检出限表"
        else:
            method = "1/2最小正值（运行时按该列计算）"
            default_replacement = np.nan
            source = "Word未提供/未匹配检出限"

        rows.append({
            "原始列名": col,
            "匹配标准名": key,
            "检出限": dl,
            "默认替代方法": method,
            "1/2检出限": default_replacement,
            "检出限来源": source,
        })
    return pd.DataFrame(rows)


# ============================================================
# 3. Excel and statistics
# ============================================================

def descriptive_stats(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy()
    x = x.apply(pd.to_numeric, errors="coerce")
    x = x.replace([np.inf, -np.inf], np.nan)

    rows = []

    for col in x.columns:
        s = x[col]
        valid = s.dropna()

        rows.append({
            "元素": col,
            "样品数": len(s),
            "有效值数量": valid.count(),
            "缺失值数量": s.isna().sum(),
            "缺失值比例": s.isna().mean(),
            "均值": valid.mean() if len(valid) else np.nan,
            "中位数": valid.median() if len(valid) else np.nan,
            "标准差": valid.std(ddof=1) if len(valid) > 1 else np.nan,
            "最小值": valid.min() if len(valid) else np.nan,
            "Q1": valid.quantile(0.25) if len(valid) else np.nan,
            "Q3": valid.quantile(0.75) if len(valid) else np.nan,
            "最大值": valid.max() if len(valid) else np.nan,
            "偏度": valid.skew() if len(valid) > 2 else np.nan,
            "峰度": valid.kurtosis() if len(valid) > 3 else np.nan,
        })

    return pd.DataFrame(rows)


def write_excel(path: Path, sheets: dict[str, pd.DataFrame]):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if df is None:
                df = pd.DataFrame()

            safe_sheet = safe_sheet_name(sheet_name)
            df.to_excel(writer, sheet_name=safe_sheet, index=False)

        workbook = writer.book

        for ws in workbook.worksheets:
            ws.freeze_panes = "A2"

            for col_cells in ws.columns:
                max_len = 10
                col_letter = col_cells[0].column_letter

                for cell in col_cells[:200]:
                    if cell.value is not None:
                        max_len = max(max_len, min(len(str(cell.value)) + 2, 50))

                ws.column_dimensions[col_letter].width = max_len


# ============================================================
# 4. Parameter dialog
# ============================================================

def ask_all_params() -> dict:
    root = tk.Tk()
    root.title("地球化学数据预处理参数设置")
    root.geometry("820x610")
    root.resizable(False, False)

    vars_dict = {
        "in_path": tk.StringVar(),
        "out_dir": tk.StringVar(),
        "series_name": tk.StringVar(),
        "elements": tk.StringVar(value=""),
        "do_log10": tk.BooleanVar(value=True),
        "do_clr": tk.BooleanVar(value=False),
        "do_ilr": tk.BooleanVar(value=False),
        "standardize": tk.StringVar(value="zscore"),
    }

    def choose_file():
        path = filedialog.askopenfilename(
            title="请选择待预处理的数据表",
            filetypes=[
                ("Excel文件", "*.xlsx;*.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*"),
            ]
        )
        if path:
            vars_dict["in_path"].set(path)
            if not vars_dict["series_name"].get().strip():
                vars_dict["series_name"].set(safe_name(Path(path).stem))

    def choose_output_dir():
        path = filedialog.askdirectory(title="请选择预处理结果保存文件夹")
        if path:
            vars_dict["out_dir"].set(path)

    def confirm():
        if not vars_dict["in_path"].get().strip():
            messagebox.showerror("错误", "请选择输入文件。")
            return

        if not vars_dict["out_dir"].get().strip():
            messagebox.showerror("错误", "请选择输出文件夹。")
            return

        scale = vars_dict["standardize"].get().strip().lower()
        if scale not in ["zscore", "z-score", "z", "minmax", "min-max", "min_max", "none", "no", "不标准化"]:
            messagebox.showerror("错误", "标准化方法只能为：zscore、minmax、none。")
            return

        root.quit()
        root.destroy()

    def cancel():
        root.quit()
        root.destroy()
        raise SystemExit("用户取消操作，程序结束。")

    row = 0

    tk.Label(root, text="输入文件：", width=18, anchor="e").grid(row=row, column=0, padx=8, pady=8, sticky="e")
    tk.Entry(root, textvariable=vars_dict["in_path"], width=78).grid(row=row, column=1, padx=4, pady=8)
    tk.Button(root, text="选择文件", command=choose_file, width=12).grid(row=row, column=2, padx=4, pady=8)

    row += 1
    tk.Label(root, text="输出文件夹：", width=18, anchor="e").grid(row=row, column=0, padx=8, pady=8, sticky="e")
    tk.Entry(root, textvariable=vars_dict["out_dir"], width=78).grid(row=row, column=1, padx=4, pady=8)
    tk.Button(root, text="选择文件夹", command=choose_output_dir, width=12).grid(row=row, column=2, padx=4, pady=8)

    row += 1
    tk.Label(root, text="结果系列名称：", width=18, anchor="e").grid(row=row, column=0, padx=8, pady=8, sticky="e")
    tk.Entry(root, textvariable=vars_dict["series_name"], width=78).grid(row=row, column=1, padx=4, pady=8)
    tk.Label(root, text="留空=输入文件名", anchor="w").grid(row=row, column=2, padx=4, pady=8)

    row += 1
    tk.Label(root, text="参与元素：", width=18, anchor="e").grid(row=row, column=0, padx=8, pady=8, sticky="e")
    tk.Entry(root, textvariable=vars_dict["elements"], width=78).grid(row=row, column=1, padx=4, pady=8)
    tk.Label(root, text="留空=全部元素", anchor="w").grid(row=row, column=2, padx=4, pady=8)

    row += 1
    tk.Label(root, text="变换方式：", width=18, anchor="e").grid(row=row, column=0, padx=8, pady=8, sticky="e")

    transform_frame = tk.Frame(root)
    transform_frame.grid(row=row, column=1, padx=4, pady=8, sticky="w")

    tk.Checkbutton(transform_frame, text="log10", variable=vars_dict["do_log10"]).grid(row=0, column=0, padx=10, sticky="w")
    tk.Checkbutton(transform_frame, text="CLR", variable=vars_dict["do_clr"]).grid(row=0, column=1, padx=10, sticky="w")
    tk.Checkbutton(transform_frame, text="ILR", variable=vars_dict["do_ilr"]).grid(row=0, column=2, padx=10, sticky="w")

    row += 1
    tk.Label(root, text="标准化：", width=18, anchor="e").grid(row=row, column=0, padx=8, pady=8, sticky="e")
    scale_frame = tk.Frame(root)
    scale_frame.grid(row=row, column=1, padx=4, pady=8, sticky="w")

    tk.Radiobutton(scale_frame, text="Z-score", variable=vars_dict["standardize"], value="zscore").grid(row=0, column=0, padx=10)
    tk.Radiobutton(scale_frame, text="MinMax", variable=vars_dict["standardize"], value="minmax").grid(row=0, column=1, padx=10)
    tk.Radiobutton(scale_frame, text="不标准化", variable=vars_dict["standardize"], value="none").grid(row=0, column=2, padx=10)

    row += 1
    info_text = (
        "说明：\n"
        "1. 默认前两列为坐标，第三列及以后为元素值。\n"
        "2. 缺失值、0值、负值：有检出限的元素按1/2检出限替代；无检出限元素按最小正值1/2替代。\n"
        "3. 若同时选择 ILR、CLR、log10，执行优先级为：ILR > CLR > log10。\n"
        "4. ILR 适合组成数据分析，会输出 ilr_1 到 ilr_D-1，因此元素列数量会减少1。\n"
        "5. CLR 输出元素数不变，但各样品行和为0。\n"
        "6. log10 适合处理右偏分布和降低极端值影响。"
    )
    tk.Label(root, text=info_text, justify="left", anchor="w").grid(row=row, column=0, columnspan=3, padx=25, pady=20, sticky="w")

    row += 1
    tk.Button(root, text="开始预处理", command=confirm, width=18, height=2).grid(row=row, column=1, padx=8, pady=15, sticky="e")
    tk.Button(root, text="取消", command=cancel, width=12, height=2).grid(row=row, column=2, padx=8, pady=15, sticky="w")

    root.mainloop()

    in_path = Path(vars_dict["in_path"].get().strip())
    out_dir = Path(vars_dict["out_dir"].get().strip())

    series_name = vars_dict["series_name"].get().strip()
    if not series_name:
        series_name = safe_name(in_path.stem)
    else:
        series_name = safe_name(series_name)

    return {
        "in_path": in_path,
        "out_dir": out_dir,
        "series_name": series_name,
        "elem_text": vars_dict["elements"].get().strip(),
        "do_log10": bool(vars_dict["do_log10"].get()),
        "do_clr": bool(vars_dict["do_clr"].get()),
        "do_ilr": bool(vars_dict["do_ilr"].get()),
        "standardize": vars_dict["standardize"].get().strip().lower(),
    }


# ============================================================
# 5. ILR, CLR and preprocessing functions
# ============================================================

def clr_transform(x: pd.DataFrame) -> pd.DataFrame:
    x_pos = x.where(x > 0, np.nan)
    ln_x = np.log(x_pos)
    row_mean_ln = ln_x.mean(axis=1, skipna=True)
    return ln_x.sub(row_mean_ln, axis=0)


def ilr_transform(x: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Apply an ILR transform using a Helmert submatrix.
    
    Args:
        x: Positive compositional data with shape n by D.
    Returns:
        An n by (D-1) ILR data frame and the D by (D-1) basis table.
    The transform is CLR(X) @ V, where V is an orthonormal basis."""

    x_pos = x.where(x > 0, np.nan)
    D = x_pos.shape[1]

    if D < 2:
        raise ValueError("ILR 至少需要2个元素列。")

    ln_x = np.log(x_pos)
    clr = ln_x.sub(ln_x.mean(axis=1, skipna=True), axis=0)

    V = np.zeros((D, D - 1), dtype=float)

    for j in range(1, D):
        V[:j, j - 1] = 1.0 / np.sqrt(j * (j + 1))
        V[j, j - 1] = -j / np.sqrt(j * (j + 1))

    ilr_values = clr.to_numpy(dtype=float) @ V

    ilr_cols = [f"ilr_{i + 1}" for i in range(D - 1)]
    ilr_df = pd.DataFrame(ilr_values, columns=ilr_cols, index=x.index)

    basis_df = pd.DataFrame(V, index=x.columns, columns=ilr_cols)
    basis_df.insert(0, "原始元素", x.columns)

    return ilr_df, basis_df.reset_index(drop=True)


def apply_preprocess(
    elements: pd.DataFrame,
    do_log10: bool = True,
    do_clr: bool = False,
    do_ilr: bool = False,
    standardize: str = "zscore",
):
    """Convert element values to numeric, coercing invalid values to NaN.
    
    Replace missing, zero and negative values with half the listed detection limit,
    or half the column minimum positive value when no limit is available. Apply the
    selected ILR, CLR or log10 transform (in that priority order), followed by the
    selected Z-score, MinMax or no standardization."""
    report_rows = []

    raw = elements.copy()
    numeric = raw.apply(pd.to_numeric, errors="coerce")
    numeric = numeric.replace([np.inf, -np.inf], np.nan)

    n_rows = len(numeric)

    # --------------------------------------------------------
    # Numeric conversion report
    # --------------------------------------------------------
    for col in numeric.columns:
        original_na = raw[col].isna().sum()
        converted_na = numeric[col].isna().sum()
        added_na = converted_na - original_na

        report_rows.append({
            "步骤": "数值转换",
            "元素": col,
            "说明": (
                "非数值内容已转换为 NaN；随后按规则替代："
                "有检出限→1/2检出限，无检出限→1/2最小正值。"
            ),
            "数量": int(max(added_na, 0)),
            "比例": float(max(added_na, 0) / n_rows) if n_rows else np.nan,
            "检出限": np.nan,
            "替代方法": "",
            "最小正值": np.nan,
            "替代值": np.nan,
        })

    processed = numeric.copy()
    replacement_rows = []

    # ========================================================
    # Replace missing, zero and negative values.
    # ========================================================
    for col in processed.columns:
        s = processed[col].copy()

        positive_values = s[s > 0].dropna()
        min_positive = float(positive_values.min()) if len(positive_values) else np.nan

        missing_count = int(s.isna().sum())
        zero_count = int((s == 0).sum(skipna=True))
        negative_count = int((s < 0).sum(skipna=True))

        replace_mask = s.isna() | (s <= 0)
        replace_count = int(replace_mask.sum())

        normalized_name = normalize_element_name(col)
        detection_limit = DETECTION_LIMITS.get(normalized_name)

        # ----------------------------------------------------
        # A. Listed detection limit: 0.5 × detection limit.
        # ----------------------------------------------------
        if detection_limit is not None:
            detection_limit = float(detection_limit)
            replacement_value = detection_limit / 2.0
            replacement_method = "1/2检出限"
            detection_source = "用户提供Word检出限表"

            s.loc[replace_mask] = replacement_value
            processed[col] = s

            explanation = (
                f"该元素匹配Word检出限：DL={detection_limit:g}；"
                f"缺失值、0值、负值统一替代为1/2 DL={replacement_value:g}。"
            )

        # ----------------------------------------------------
        # B. No listed detection limit: 0.5 × minimum positive value.
        # ----------------------------------------------------
        else:
            detection_limit = np.nan
            detection_source = "Word未提供/未匹配"

            if len(positive_values) == 0:
                replacement_value = np.nan
                replacement_method = "无法替代"
                processed[col] = s
                explanation = (
                    "Word表中无该元素检出限，且该列没有任何正值，"
                    "无法使用1/2最小正值替代；建议检查或剔除该变量。"
                )
            else:
                replacement_value = min_positive / 2.0
                replacement_method = "1/2最小正值"
                s.loc[replace_mask] = replacement_value
                processed[col] = s
                explanation = (
                    "Word表中无该元素检出限，"
                    f"因此缺失值、0值、负值使用该列最小正值"
                    f"({min_positive:g})的1/2，即{replacement_value:g}替代。"
                )

        replacement_rows.append({
            "步骤": "缺失值与非正值替代",
            "元素": col,
            "匹配标准名": normalized_name,
            "检出限": detection_limit,
            "检出限来源": detection_source,
            "替代方法": replacement_method,
            "说明": explanation,
            "缺失值数量": missing_count,
            "0值数量": zero_count,
            "负值数量": negative_count,
            "总替代数量": replace_count,
            "总替代比例": float(replace_count / n_rows) if n_rows else np.nan,
            "最小正值": min_positive,
            "最终替代值": replacement_value,
        })

        report_rows.append({
            "步骤": "缺失值与非正值替代",
            "元素": col,
            "说明": explanation,
            "数量": replace_count,
            "比例": float(replace_count / n_rows) if n_rows else np.nan,
            "检出限": detection_limit,
            "替代方法": replacement_method,
            "最小正值": min_positive,
            "替代值": replacement_value,
        })

    replacement_report = pd.DataFrame(replacement_rows)
    detection_limit_report = build_detection_limit_table(list(processed.columns))

    # ========================================================
    # ILR / CLR / log10
    # ========================================================
    ilr_basis_df = pd.DataFrame()

    if do_ilr:
        non_positive_count = ((processed <= 0) & processed.notna()).sum()

        for col in processed.columns:
            report_rows.append({
                "步骤": "ILR前正值检查",
                "元素": col,
                "说明": "ILR要求所有参与元素值>0。",
                "数量": int(non_positive_count[col]),
                "比例": float(non_positive_count[col] / n_rows) if n_rows else np.nan,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

        processed, ilr_basis_df = ilr_transform(processed)

        report_rows.append({
            "步骤": "ILR",
            "元素": "全部参与元素",
            "说明": "已执行 ILR 转换；ILR优先级高于CLR和log10，输出维度为D-1。",
            "数量": len(processed),
            "比例": np.nan,
            "检出限": np.nan,
            "替代方法": "",
            "最小正值": np.nan,
            "替代值": np.nan,
        })

    elif do_clr:
        non_positive_count = ((processed <= 0) & processed.notna()).sum()

        for col in processed.columns:
            report_rows.append({
                "步骤": "CLR前正值检查",
                "元素": col,
                "说明": "CLR要求所有参与元素值>0。",
                "数量": int(non_positive_count[col]),
                "比例": float(non_positive_count[col] / n_rows) if n_rows else np.nan,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

        processed = clr_transform(processed)

        report_rows.append({
            "步骤": "CLR",
            "元素": "全部参与元素",
            "说明": "已执行 CLR：ln(x_i / g(x))；CLR优先级高于普通log10。",
            "数量": len(processed),
            "比例": np.nan,
            "检出限": np.nan,
            "替代方法": "",
            "最小正值": np.nan,
            "替代值": np.nan,
        })

    elif do_log10:
        non_positive_count = ((processed <= 0) & processed.notna()).sum()

        for col in processed.columns:
            report_rows.append({
                "步骤": "log10前正值检查",
                "元素": col,
                "说明": "log10要求元素值>0。",
                "数量": int(non_positive_count[col]),
                "比例": float(non_positive_count[col] / n_rows) if n_rows else np.nan,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

        processed = processed.where(processed > 0, np.nan)
        processed = np.log10(processed)

        report_rows.append({
            "步骤": "log10",
            "元素": "全部参与元素",
            "说明": "已对完成检出限/最小正值替代后的数据执行 log10 转换。",
            "数量": len(processed),
            "比例": np.nan,
            "检出限": np.nan,
            "替代方法": "",
            "最小正值": np.nan,
            "替代值": np.nan,
        })

    else:
        report_rows.append({
            "步骤": "log10/CLR/ILR",
            "元素": "全部参与元素",
            "说明": (
                "未执行 log10、CLR 或 ILR；"
                "但已完成缺失值、0值、负值替代："
                "有检出限按1/2检出限，无检出限按1/2最小正值。"
            ),
            "数量": len(processed),
            "比例": np.nan,
            "检出限": np.nan,
            "替代方法": "",
            "最小正值": np.nan,
            "替代值": np.nan,
        })

    # ========================================================
    # Standardization
    # ========================================================
    standardize = str(standardize).strip().lower()

    if standardize in ["zscore", "z-score", "z"]:
        means = processed.mean(axis=0, skipna=True)
        stds = processed.std(axis=0, ddof=1, skipna=True)

        zero_std_cols = stds[stds == 0].index.tolist()

        for col in processed.columns:
            report_rows.append({
                "步骤": "Z-score标准化",
                "元素": col,
                "说明": f"mean={means[col]}, std={stds[col]}",
                "数量": int(processed[col].notna().sum()),
                "比例": np.nan,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

        processed = (processed - means) / stds.replace(0, np.nan)

        for col in zero_std_cols:
            report_rows.append({
                "步骤": "Z-score异常",
                "元素": col,
                "说明": "该列标准差为0，标准化后结果为 NaN；建议剔除该变量。",
                "数量": int(len(processed)),
                "比例": 1.0,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

    elif standardize in ["minmax", "min-max", "min_max"]:
        mins = processed.min(axis=0, skipna=True)
        maxs = processed.max(axis=0, skipna=True)
        ranges = maxs - mins

        zero_range_cols = ranges[ranges == 0].index.tolist()

        for col in processed.columns:
            report_rows.append({
                "步骤": "MinMax标准化",
                "元素": col,
                "说明": f"min={mins[col]}, max={maxs[col]}",
                "数量": int(processed[col].notna().sum()),
                "比例": np.nan,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

        processed = (processed - mins) / ranges.replace(0, np.nan)

        for col in zero_range_cols:
            report_rows.append({
                "步骤": "MinMax异常",
                "元素": col,
                "说明": "该列最大值等于最小值，标准化后结果为 NaN；建议剔除该变量。",
                "数量": int(len(processed)),
                "比例": 1.0,
                "检出限": np.nan,
                "替代方法": "",
                "最小正值": np.nan,
                "替代值": np.nan,
            })

    elif standardize in ["none", "no", "不标准化"]:
        report_rows.append({
            "步骤": "标准化",
            "元素": "全部变量",
            "说明": "未执行标准化",
            "数量": len(processed),
            "比例": np.nan,
            "检出限": np.nan,
            "替代方法": "",
            "最小正值": np.nan,
            "替代值": np.nan,
        })

    else:
        raise ValueError("标准化方法只能为：zscore、minmax、none。")

    step_report = pd.DataFrame(report_rows)
    step_report.attrs["replacement_report"] = replacement_report
    step_report.attrs["detection_limit_report"] = detection_limit_report
    step_report.attrs["ilr_basis_df"] = ilr_basis_df

    return processed, step_report


# ============================================================
# 6. Main program
# ============================================================

def main():
    try:
        params = ask_all_params()

        in_path = params["in_path"]
        out_dir = params["out_dir"]
        series_name = params["series_name"]
        elem_text = params["elem_text"]
        do_log10 = params["do_log10"]
        do_clr = params["do_clr"]
        do_ilr = params["do_ilr"]
        scale = params["standardize"]

        df = read_table(in_path)
        coords, elements, coord_cols, element_cols = split_geochem_table(df)

        selected = parse_columns(elem_text, element_cols, default_all=True)
        elements_selected = elements[selected].copy()

        processed, step_report = apply_preprocess(
            elements_selected,
            do_log10=do_log10,
            do_clr=do_clr,
            do_ilr=do_ilr,
            standardize=scale
        )

        result = pd.concat(
            [
                coords.reset_index(drop=True),
                processed.reset_index(drop=True)
            ],
            axis=1
        )

        valid_mask = processed.notna().all(axis=1)
        model_ready = result.loc[valid_mask].reset_index(drop=True)

        tag = now_tag()

        transform_name = "ILR" if do_ilr else ("CLR" if do_clr else ("log10" if do_log10 else "raw"))
        full_path = out_dir / f"{series_name}_{transform_name}_{scale}_预处理完整结果_{tag}.xlsx"
        ready_path = out_dir / f"{series_name}_{transform_name}_{scale}_建模可用数据_{tag}.xlsx"
        report_path = out_dir / f"{series_name}_{transform_name}_{scale}_预处理报告_{tag}.xlsx"

        result.to_excel(full_path, index=False, engine="openpyxl")
        model_ready.to_excel(ready_path, index=False, engine="openpyxl")

        summary = pd.DataFrame([
            {"项目": "输入文件", "结果": str(in_path)},
            {"项目": "输出文件夹", "结果": str(out_dir)},
            {"项目": "输出文件系列名称", "结果": series_name},
            {"项目": "完整结果文件", "结果": str(full_path)},
            {"项目": "建模可用数据文件", "结果": str(ready_path)},
            {"项目": "报告文件", "结果": str(report_path)},
            {"项目": "数据结构", "结果": "默认前两列为坐标，第三列以后为元素"},
            {"项目": "坐标列", "结果": "、".join(coord_cols)},
            {"项目": "原始元素数量", "结果": len(element_cols)},
            {"项目": "参与元素数量", "结果": len(selected)},
            {"项目": "参与元素", "结果": "、".join(selected)},
            {"项目": "log10", "结果": "是" if do_log10 else "否"},
            {"项目": "CLR", "结果": "是" if do_clr else "否"},
            {"项目": "ILR", "结果": "是" if do_ilr else "否"},
            {"项目": "实际执行变换", "结果": transform_name},
            {"项目": "标准化", "结果": scale},
            {"项目": "原始样品数", "结果": len(df)},
            {"项目": "建模可用样品数", "结果": int(valid_mask.sum())},
            {"项目": "剔除样品数", "结果": int((~valid_mask).sum())},
            {"项目": "缺失值与非正值处理", "结果": "有Word检出限的元素：缺失值、0值、负值按1/2检出限替代；无检出限元素按该列最小正值1/2替代。"},
            {"项目": "检出限来源", "结果": "用户提供的Word《各元素检出限》；未提供检出限的元素自动采用1/2最小正值。"},
            {"项目": "检出限单位说明", "结果": "检测限按Word中的科学计数值原样录入；程序不自动进行ppm、ppb、wt.%或质量分数换算，输入数据须与检测限单位体系一致。"},
            {"项目": "变换优先级", "结果": "ILR > CLR > log10"},
            {"项目": "ILR说明", "结果": "ILR输出维度为D-1，列名为ilr_1至ilr_D-1。"},
        ])

        invalid_rows = result.loc[~valid_mask].copy()
        if len(invalid_rows):
            invalid_rows.insert(0, "原始DataFrame索引", result.index[~valid_mask])

        replacement_report = step_report.attrs.get("replacement_report", pd.DataFrame())
        detection_limit_report = step_report.attrs.get("detection_limit_report", pd.DataFrame())
        ilr_basis_df = step_report.attrs.get("ilr_basis_df", pd.DataFrame())

        write_excel(report_path, {
            "预处理汇总": summary,
            "检出限与替代规则": detection_limit_report,
            "替代值明细": replacement_report,
            "步骤报告": step_report,
            "ILR基矩阵": ilr_basis_df,
            "预处理前统计": descriptive_stats(elements_selected),
            "预处理后统计": descriptive_stats(processed),
            "未进入建模样品": invalid_rows,
        })

        print("完成。")
        print(f"结果系列名称：{series_name}")
        print(f"完整结果：{full_path}")
        print(f"建模可用数据：{ready_path}")
        print(f"报告：{report_path}")

        messagebox.showinfo(
            "完成",
            f"预处理完成！\n\n"
            f"完整结果：\n{full_path}\n\n"
            f"建模可用数据：\n{ready_path}\n\n"
            f"报告：\n{report_path}"
        )

        open_output_folder(out_dir)

    except SystemExit as e:
        print(str(e))

    except Exception as e:
        print("程序运行出错：", e)
        messagebox.showerror("程序运行出错", f"错误信息：\n{e}")


if __name__ == "__main__":
    main()