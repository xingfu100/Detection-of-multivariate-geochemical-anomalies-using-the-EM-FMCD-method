# coding=utf-8

"""Fit diagonal-covariance Gaussian mixtures using EM and multi-seed AIC.

Input: Excel or CSV with two coordinate columns followed by clustering variables.
Options: K range, random seeds, final K values, and output directory.
Output: AIC plot and Excel workbooks with cluster and model statistics.
Run: Execute this script and follow the file and parameter dialogs."""

from __future__ import annotations

import os
import sys
import time
import traceback
import subprocess
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

from sklearn.mixture import GaussianMixture


GMM_MAX_ITER = 1500
GMM_TOL = 1e-6


GMM_INIT_PARAMS = "random_from_data"

GMM_COVARIANCE_TYPE = "diag"
GMM_REG_COVAR = 1e-6


GMM_N_INIT = 1

DEFAULT_SEED_LIST = "1,7,13,21,42,66,88,100,123,2024"


def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def safe_name(name: str) -> str:
    name = str(name)
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    for ch in invalid_chars:
        name = name.replace(ch, "_")
    return name.strip()


def safe_sheet_name(name: str) -> str:
    name = str(name)
    for ch in ["\\", "/", "?", "*", "[", "]", ":"]:
        name = name.replace(ch, "_")
    return name[:31]


def make_unique_columns(columns) -> list[str]:
    seen = {}
    out = []

    for col in columns:
        c = str(col).strip()
        if c == "" or c.lower() == "nan":
            c = "未命名列"

        if c not in seen:
            seen[c] = 0
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c}_重复{seen[c]}")

    return out


def read_table(file_path: str | Path) -> pd.DataFrame:
    file_path = str(file_path)
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xlsx":
        return pd.read_excel(file_path, engine="openpyxl")
    if ext == ".xls":
        return pd.read_excel(file_path, engine="xlrd")
    if ext == ".csv":
        return pd.read_csv(file_path, encoding="utf-8-sig")

    raise ValueError("仅支持 .xlsx、.xls、.csv 文件。")


def write_excel(path: str | Path, sheets: dict[str, pd.DataFrame]) -> None:
    path = Path(path)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_sheet = safe_sheet_name(sheet_name)

            if df is None:
                df = pd.DataFrame()

            if not isinstance(df, pd.DataFrame):
                df = pd.DataFrame(df)

            df.to_excel(writer, sheet_name=safe_sheet, index=False)

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"

            for col_cells in ws.columns:
                max_len = 10
                col_letter = col_cells[0].column_letter

                for cell in col_cells[:100]:
                    if cell.value is not None:
                        max_len = max(max_len, min(len(str(cell.value)) + 2, 45))

                ws.column_dimensions[col_letter].width = max_len




def open_file(path: str | Path) -> None:
    path = Path(path)

    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        print(f"无法自动打开文件：{path}")
        print(e)


def open_output_folder(path: str | Path) -> None:
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)

    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        print(f"无法自动打开文件夹：{path}")
        print(e)


def ask_file_path(title: str) -> Path:
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel文件", "*.xlsx;*.xls"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*"),
        ],
    )

    if not file_path:
        raise RuntimeError("未选择输入文件。")

    return Path(file_path)


def ask_text(title: str, prompt: str, initialvalue: str = "") -> str:
    root = tk.Tk()
    root.withdraw()

    value = simpledialog.askstring(
        title=title,
        prompt=prompt,
        initialvalue=initialvalue,
    )

    if value is None:
        raise RuntimeError(f"用户取消输入：{title}")

    return str(value).strip()


def ask_output_dir(default_folder_name: str = "EM聚类AIC多种子分析结果") -> Path:
    root = tk.Tk()
    root.withdraw()

    parent_folder = filedialog.askdirectory(
        title="请选择结果文件夹的父级保存位置"
    )

    if not parent_folder:
        raise RuntimeError("未选择输出父级文件夹。")

    default_name = f"{default_folder_name}_{now_tag()}"

    folder_name = simpledialog.askstring(
        title="新建结果文件夹名称",
        prompt=(
            "请输入本次结果的新文件夹名称。\n\n"
            "所有AIC曲线和Excel结果都会保存在该文件夹中。\n"
            "如果直接回车，将使用默认名称。"
        ),
        initialvalue=default_name,
    )

    if folder_name is None:
        raise RuntimeError("用户取消输入结果文件夹名称。")

    folder_name = safe_name(folder_name.strip()) if folder_name.strip() else default_name

    out_dir = Path(parent_folder) / folder_name

    if out_dir.exists():
        out_dir = Path(parent_folder) / f"{folder_name}_{now_tag()}"
        messagebox.showinfo(
            "文件夹已存在",
            f"你输入的文件夹名称已存在。\n\n程序将自动新建文件夹：\n{out_dir}"
        )

    out_dir.mkdir(parents=True, exist_ok=False)

    return out_dir


def parse_range_text(text: str, name: str, min_allowed: int = 1) -> tuple[int, int]:
    if text is None or str(text).strip() == "":
        raise ValueError(f"{name}不能为空。")

    s = str(text).strip()
    s = s.replace("，", ",").replace("；", ",").replace(";", ",")
    s = s.replace("－", "-").replace("—", "-").replace("~", "-")

    if "-" in s and "," not in s:
        parts = [x.strip() for x in s.split("-") if x.strip()]
    elif "," in s:
        parts = [x.strip() for x in s.split(",") if x.strip()]
    else:
        parts = [x.strip() for x in s.split() if x.strip()]

    if len(parts) == 1:
        a = b = int(parts[0])
    elif len(parts) == 2:
        a = int(parts[0])
        b = int(parts[1])
    else:
        raise ValueError(f"{name}格式错误。示例：2-12 或 6。")

    if a > b:
        raise ValueError(f"{name}最小值不能大于最大值。")

    if a < min_allowed:
        raise ValueError(f"{name}不能小于 {min_allowed}。")

    return a, b


def ask_k_range(default_min: int = 2, default_max: int = 10) -> tuple[int, int]:
    while True:
        text = ask_text(
            "AIC曲线K值范围",
            (
                "请输入用于绘制AIC曲线的K值范围。\n\n"
                "示例：2-12\n"
                "示例：3-11\n"
                "示例：6 表示只计算K=6\n\n"
                "建议先用 2-12 或 3-11。"
            ),
            f"{default_min}-{default_max}",
        )

        try:
            return parse_range_text(text, "K值范围", min_allowed=1)
        except Exception as e:
            messagebox.showerror("输入错误", str(e))


def parse_k_values(text: str, min_value: int = 1, max_value: int | None = None) -> list[int]:
    if text is None or str(text).strip() == "":
        raise ValueError("K值不能为空。")

    s = str(text).strip()
    s = s.replace("，", ",").replace("；", ",").replace(";", ",")
    s = s.replace("－", "-").replace("—", "-").replace("~", "-")

    values: list[int] = []

    if "-" in s and "," not in s:
        parts = [x.strip() for x in s.split("-") if x.strip()]
        if len(parts) != 2:
            raise ValueError("K范围格式错误。示例：5-7。")
        k_min = int(parts[0])
        k_max = int(parts[1])
        if k_min > k_max:
            raise ValueError("K最小值不能大于K最大值。")
        values = list(range(k_min, k_max + 1))
    else:
        parts = [x.strip() for x in s.replace(",", " ").split() if x.strip()]
        values = [int(x) for x in parts]

    values = sorted(set(values))

    for k in values:
        if k < min_value:
            raise ValueError(f"K值不能小于 {min_value}。")
        if max_value is not None and k > max_value:
            raise ValueError(f"K值不能大于 {max_value}。")

    return values


def ask_final_k_values(max_k: int, default_text: str = "6") -> list[int]:
    while True:
        text = ask_text(
            "最终EM聚类K值",
            (
                "AIC曲线已经生成并打开。\n\n"
                "请根据AIC曲线和地质解释需要，输入最终用于EM聚类的K值。\n\n"
                "支持单值或范围：\n"
                "例如：6\n"
                "例如：5-7\n"
                "例如：5,6,7\n\n"
                f"当前最大允许K值：{max_k}"
            ),
            default_text,
        )

        try:
            return parse_k_values(text, min_value=1, max_value=max_k)
        except Exception as e:
            messagebox.showerror("K值输入错误", str(e))


def parse_seed_list(text: str) -> list[int]:
    """Parse a list, range, or random set of seeds."""
    if text is None or str(text).strip() == "":
        text = DEFAULT_SEED_LIST

    s = str(text).strip()
    s_lower = s.lower()

    if s_lower in ["none", "random", "随机"]:
        seeds = np.random.SeedSequence().generate_state(20, dtype=np.uint32)
        return sorted(set(int(x) for x in seeds))

    s = s.replace("，", ",").replace("；", ",").replace(";", ",")
    s = s.replace("－", "-").replace("—", "-").replace("~", "-")

    values: list[int] = []

    if "-" in s and "," not in s:
        parts = [x.strip() for x in s.split("-") if x.strip()]
        if len(parts) != 2:
            raise ValueError("随机种子范围格式错误。示例：1-20。")
        a = int(parts[0])
        b = int(parts[1])
        if a > b:
            raise ValueError("随机种子范围最小值不能大于最大值。")
        values = list(range(a, b + 1))
    else:
        parts = [x.strip() for x in s.replace(",", " ").split() if x.strip()]
        values = [int(x) for x in parts]

    values = sorted(set(values))

    if len(values) == 0:
        raise ValueError("随机种子列表为空。")

    for v in values:
        if v < 0:
            raise ValueError("随机种子必须为非负整数。")

    return values


def ask_seed_list() -> list[int]:
    while True:
        text = ask_text(
            "多随机种子设置",
            (
                "请输入多个随机种子。\n\n"
                "示例：1,7,13,21,42,66,88,100,123,2024\n"
                "示例：1-30 表示使用1到30共30个种子\n"
                "输入 None 或 random：随机生成20个真实整数种子\n\n"
                "程序会对每个K分别使用这些随机种子重复拟合，"
                "并自动选择AIC最小的模型作为该K最佳结果。"
            ),
            DEFAULT_SEED_LIST,
        )

        try:
            seeds = parse_seed_list(text)
            messagebox.showinfo(
                "随机种子确认",
                f"本次将使用 {len(seeds)} 个随机种子：\n{seeds}"
            )
            return seeds
        except Exception as e:
            messagebox.showerror("随机种子输入错误", str(e))


def split_geochem_table(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, list[str], list[str]]:
    df = df.copy()
    df.columns = make_unique_columns(df.columns)

    if df.shape[1] < 3:
        raise ValueError("数据表至少需要3列：前两列为坐标，第三列以后为元素。")

    coord_cols = df.columns[:2].tolist()
    element_cols = df.columns[2:].tolist()

    coords = df[coord_cols].copy()
    elements = df[element_cols].copy()

    return coords, elements, coord_cols, element_cols


def descriptive_stats(df: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for col in df.columns:
        s = pd.to_numeric(df[col], errors="coerce")
        s = s.replace([np.inf, -np.inf], np.nan).dropna()

        if len(s) == 0:
            rows.append({
                "元素": col,
                "有效样品数": 0,
                "均值": np.nan,
                "中位数": np.nan,
                "标准差": np.nan,
                "最小值": np.nan,
                "P05": np.nan,
                "P25": np.nan,
                "P75": np.nan,
                "P95": np.nan,
                "最大值": np.nan,
                "偏度": np.nan,
                "峰度": np.nan,
            })
            continue

        rows.append({
            "元素": col,
            "有效样品数": len(s),
            "均值": s.mean(),
            "中位数": s.median(),
            "标准差": s.std(ddof=1),
            "最小值": s.min(),
            "P05": s.quantile(0.05),
            "P25": s.quantile(0.25),
            "P75": s.quantile(0.75),
            "P95": s.quantile(0.95),
            "最大值": s.max(),
            "偏度": s.skew(),
            "峰度": s.kurt(),
        })

    return pd.DataFrame(rows)


def gmm_parameter_count(k: int, n_vars: int, covariance_type: str = "diag") -> int:
    if covariance_type == "diag":
        return (k - 1) + 2 * k * n_vars

    if covariance_type == "full":
        return (k - 1) + k * n_vars + k * n_vars * (n_vars + 1) // 2

    raise ValueError(f"暂未实现 covariance_type={covariance_type} 的参数数量计算。")


def posterior_entropy(probabilities: np.ndarray) -> np.ndarray:
    eps = 1e-12
    p = np.clip(probabilities, eps, 1.0)
    return -np.sum(p * np.log(p), axis=1)


def fit_gmm_diag_local(X: np.ndarray, k: int, seed: int) -> GaussianMixture:
    model = GaussianMixture(
        n_components=k,
        covariance_type=GMM_COVARIANCE_TYPE,
        init_params=GMM_INIT_PARAMS,
        max_iter=GMM_MAX_ITER,
        tol=GMM_TOL,
        reg_covar=GMM_REG_COVAR,
        n_init=GMM_N_INIT,
        random_state=int(seed),
    )

    model.fit(X)
    return model


def fit_best_gmm_by_seeds(
    X: np.ndarray,
    k: int,
    seed_list: list[int],
) -> tuple[GaussianMixture, pd.DataFrame]:
    rows = []
    best_model = None
    best_aic = np.inf

    n_samples, n_vars = X.shape

    # Try each seed for the current number of mixture components.
    for seed in seed_list:
        print(f"正在拟合：K={k}，seed={seed}")

        model = fit_gmm_diag_local(X, k, seed)

        log_likelihood_total = float(model.score(X) * n_samples)
        aic_value = float(model.aic(X))
        bic_value = float(model.bic(X))
        n_params = gmm_parameter_count(k, n_vars, GMM_COVARIANCE_TYPE)

        rows.append({
            "K": k,
            "随机种子": int(seed),
            "AIC": aic_value,
            "BIC": bic_value,
            "log_likelihood_total": log_likelihood_total,
            "平均log_likelihood": float(model.score(X)),
            "理论自由参数数_diag": n_params,
            "样品数": n_samples,
            "变量数": n_vars,
            "是否收敛": "是" if model.converged_ else "否",
            "迭代次数": model.n_iter_,
            "max_iter": GMM_MAX_ITER,
            "tol": GMM_TOL,
            "init_params": GMM_INIT_PARAMS,
            "covariance_type": GMM_COVARIANCE_TYPE,
            "reg_covar": GMM_REG_COVAR,
            "n_init": GMM_N_INIT,
        })

        if aic_value < best_aic:
            best_aic = aic_value
            best_model = model

    if best_model is None:
        raise RuntimeError(f"K={k} 未得到有效模型。")

    trial_df = pd.DataFrame(rows)
    trial_df["该K内AIC排序"] = trial_df["AIC"].rank(method="min", ascending=True).astype(int)
    trial_df["是否该K最佳seed"] = np.where(trial_df["该K内AIC排序"] == 1, "是", "否")

    return best_model, trial_df


def build_aic_curve_table_multi_seed(
    X: np.ndarray,
    k_min: int,
    k_max: int,
    seed_list: list[int],
) -> tuple[pd.DataFrame, pd.DataFrame, dict[int, GaussianMixture]]:
    all_trials = []
    best_models: dict[int, GaussianMixture] = {}

    n_samples, n_vars = X.shape

    for k in range(k_min, k_max + 1):
        if k >= n_samples:
            continue

        best_model, trial_df = fit_best_gmm_by_seeds(X, k, seed_list)

        best_models[k] = best_model
        all_trials.append(trial_df)

    all_trial_df = pd.concat(all_trials, ignore_index=True)

    best_rows = (
        all_trial_df.sort_values(["K", "AIC"], ascending=[True, True])
        .groupby("K", as_index=False)
        .first()
    )

    curve_df = best_rows.copy()
    curve_df = curve_df.rename(columns={
        "随机种子": "该K最佳随机种子",
        "AIC": "最佳AIC",
        "BIC": "最佳BIC",
        "log_likelihood_total": "最佳log_likelihood_total",
        "平均log_likelihood": "最佳平均log_likelihood",
        "是否收敛": "最佳模型是否收敛",
        "迭代次数": "最佳模型迭代次数",
    })

    curve_df["AIC"] = curve_df["最佳AIC"]
    curve_df["BIC"] = curve_df["最佳BIC"]

    curve_df = curve_df.sort_values("K").reset_index(drop=True)

    if not curve_df.empty:
        curve_df["delta_AIC"] = curve_df["AIC"] - curve_df["AIC"].min()
        curve_df["AIC相对似然"] = np.exp(-0.5 * curve_df["delta_AIC"])
        curve_df["AIC排序"] = curve_df["AIC"].rank(method="min", ascending=True).astype(int)
        curve_df["相邻K_AIC变化量"] = curve_df["AIC"].diff()
        curve_df["相邻K_AIC下降量"] = -curve_df["相邻K_AIC变化量"]

    return curve_df, all_trial_df, best_models


def plot_aic_only(curve_df: pd.DataFrame, output_prefix: Path) -> Path:
    out_png = output_prefix.with_name(output_prefix.name + "_AIC曲线_多种子最佳.png")

    fig, ax = plt.subplots(figsize=(8.5, 5.2), dpi=300)

    ax.plot(
        curve_df["K"],
        curve_df["AIC"],
        marker="o",
        linewidth=1.8,
    )

    ax.set_xlabel("Number of Components (K)")
    ax.set_ylabel("Best AIC across seeds")
    ax.set_title("AIC Curve for EM-GMM: Best Model across Multiple Seeds")
    ax.set_xticks(curve_df["K"].astype(int).tolist())
    ax.grid(True, linestyle="--", alpha=0.35)

    best_idx = curve_df["AIC"].idxmin()
    best_k = int(curve_df.loc[best_idx, "K"])
    best_aic = float(curve_df.loc[best_idx, "AIC"])
    best_seed = int(curve_df.loc[best_idx, "该K最佳随机种子"])

    ax.scatter([best_k], [best_aic], s=55, zorder=5)
    ax.annotate(
        f"Min AIC: K={best_k}, seed={best_seed}",
        xy=(best_k, best_aic),
        xytext=(best_k, best_aic + abs(best_aic) * 0.03 if best_aic != 0 else best_aic + 1),
        arrowprops=dict(arrowstyle="->", lw=0.8),
        fontsize=9,
    )

    fig.tight_layout()
    fig.savefig(out_png)
    plt.close(fig)

    return out_png


def build_cluster_result(
    valid_original: pd.DataFrame,
    coord_cols: list[str],
    model: GaussianMixture,
    X: np.ndarray,
) -> pd.DataFrame:
    labels = model.predict(X) + 1
    probabilities = model.predict_proba(X)

    max_prob = probabilities.max(axis=1)
    uncertainty = 1.0 - max_prob
    entropy = posterior_entropy(probabilities)

    result = valid_original.copy()
    result["EM类别"] = labels
    result["EM最大后验概率"] = max_prob
    result["EM分类不确定性"] = uncertainty
    result["EM概率熵"] = entropy

    for i in range(probabilities.shape[1]):
        result[f"类别{i + 1}_后验概率"] = probabilities[:, i]

    count_map = pd.Series(labels).value_counts().sort_index().to_dict()
    total_n = len(labels)
    result["所属类别样品数"] = result["EM类别"].map(count_map)
    result["所属类别样品比例"] = result["所属类别样品数"] / total_n

    em_cols = [
        "EM类别",
        "所属类别样品数",
        "所属类别样品比例",
        "EM最大后验概率",
        "EM分类不确定性",
        "EM概率熵"
    ]
    prob_cols = [f"类别{i + 1}_后验概率" for i in range(probabilities.shape[1])]

    other_cols = [c for c in result.columns if c not in coord_cols + em_cols + prob_cols]
    ordered_cols = coord_cols + em_cols + prob_cols + other_cols

    return result[ordered_cols]


def build_class_count_table(result_df: pd.DataFrame) -> pd.DataFrame:
    total_n = len(result_df)

    out = (
        result_df.groupby("EM类别")
        .agg(
            样品数=("EM类别", "size"),
            平均最大后验概率=("EM最大后验概率", "mean"),
            中位最大后验概率=("EM最大后验概率", "median"),
            平均分类不确定性=("EM分类不确定性", "mean"),
            平均概率熵=("EM概率熵", "mean"),
        )
        .reset_index()
    )

    out["样品比例"] = out["样品数"] / total_n
    out["样品比例_%"] = out["样品比例"] * 100

    out = out[
        [
            "EM类别",
            "样品数",
            "样品比例",
            "样品比例_%",
            "平均最大后验概率",
            "中位最大后验概率",
            "平均分类不确定性",
            "平均概率熵",
        ]
    ]

    return out


def build_concise_class_summary(
    result_df: pd.DataFrame,
    weights_df: pd.DataFrame | None = None
) -> pd.DataFrame:
    count_df = build_class_count_table(result_df)

    if weights_df is not None and not weights_df.empty:
        count_df = pd.merge(count_df, weights_df, on="EM类别", how="left")

    return count_df


def cluster_summary_detail(result_with_elements: pd.DataFrame, element_cols: list[str]) -> pd.DataFrame:
    rows = []

    for label, g in result_with_elements.groupby("EM类别"):
        row = {
            "EM类别": label,
            "样品数": len(g),
            "样品比例": len(g) / len(result_with_elements),
            "平均最大后验概率": g["EM最大后验概率"].mean(),
            "最小最大后验概率": g["EM最大后验概率"].min(),
            "中位最大后验概率": g["EM最大后验概率"].median(),
            "平均分类不确定性": g["EM分类不确定性"].mean(),
            "平均概率熵": g["EM概率熵"].mean(),
        }

        for col in element_cols:
            s = g[col]
            row[f"{col}_均值"] = s.mean()
            row[f"{col}_中位数"] = s.median()
            row[f"{col}_标准差"] = s.std(ddof=1)
            row[f"{col}_P05"] = s.quantile(0.05)
            row[f"{col}_P95"] = s.quantile(0.95)

        rows.append(row)

    return pd.DataFrame(rows)


def posterior_probability_summary(result_df: pd.DataFrame, k: int) -> pd.DataFrame:
    rows = []

    for cls in range(1, k + 1):
        col = f"类别{cls}_后验概率"
        if col not in result_df.columns:
            continue

        s = result_df[col]
        rows.append({
            "类别": cls,
            "该类后验概率均值": s.mean(),
            "该类后验概率中位数": s.median(),
            "该类后验概率P05": s.quantile(0.05),
            "该类后验概率P95": s.quantile(0.95),
            "该类后验概率最大值": s.max(),
        })

    return pd.DataFrame(rows)


def class_mean_distance(means_df: pd.DataFrame, element_cols: list[str]) -> pd.DataFrame:
    rows = []

    for i in range(len(means_df)):
        for j in range(i + 1, len(means_df)):
            ci = int(means_df.loc[i, "EM类别"])
            cj = int(means_df.loc[j, "EM类别"])

            vi = means_df.loc[i, element_cols].to_numpy(dtype=float)
            vj = means_df.loc[j, element_cols].to_numpy(dtype=float)

            rows.append({
                "类别1": ci,
                "类别2": cj,
                "均值向量欧氏距离": float(np.linalg.norm(vi - vj)),
            })

    return pd.DataFrame(rows)


def build_seed_record_df(seed_list: list[int]) -> pd.DataFrame:
    return pd.DataFrame([{
        "随机种子数量": len(seed_list),
        "随机种子列表": "、".join(map(str, seed_list)),
        "复现说明": "复现实验时使用同一随机种子列表、同一K范围、同一GMM参数。"
    }])


def build_model_tables(
    model: GaussianMixture,
    X: np.ndarray,
    selected: list[str],
    final_k: int,
    best_seed: int,
    seed_list: list[int],
    aic_png: Path,
) -> dict[str, pd.DataFrame]:
    n_samples, n_vars = X.shape
    log_likelihood_total = float(model.score(X) * n_samples)
    aic_value = float(model.aic(X))
    bic_value = float(model.bic(X))
    n_params = gmm_parameter_count(final_k, n_vars, GMM_COVARIANCE_TYPE)

    model_info = {
        "最终K值": final_k,
        "该K最佳随机种子": best_seed,
        "随机种子试验数量": len(seed_list),
        "随机种子列表": "、".join(map(str, seed_list)),
        "复现说明": f"复现实验时使用 K={final_k}，RANDOM_STATE={best_seed}",
        "样品数": n_samples,
        "变量数": n_vars,
        "参与元素数": len(selected),
        "参与元素": "、".join(selected),
        "协方差类型": GMM_COVARIANCE_TYPE,
        "初始化方式": GMM_INIT_PARAMS,
        "最大迭代次数": GMM_MAX_ITER,
        "收敛容差": GMM_TOL,
        "正则化参数reg_covar": GMM_REG_COVAR,
        "初始化次数n_init": GMM_N_INIT,
        "是否收敛": "是" if model.converged_ else "否",
        "迭代次数": model.n_iter_,
        "log_likelihood_total": log_likelihood_total,
        "平均log_likelihood": float(model.score(X)),
        "理论自由参数数_diag": n_params,
        "AIC": aic_value,
        "BIC": bic_value,
        "AIC计算说明": "AIC = 2p - 2ln(L)，p为模型自由参数数，L为似然函数。",
        "AIC曲线图": str(aic_png),
    }

    model_info_df = pd.DataFrame([{"参数": k, "取值": v} for k, v in model_info.items()])

    weights = pd.DataFrame({
        "EM类别": np.arange(1, final_k + 1),
        "混合权重": model.weights_,
    })

    means = pd.DataFrame(model.means_, columns=selected)
    means.insert(0, "EM类别", np.arange(1, final_k + 1))

    variances = pd.DataFrame(model.covariances_, columns=[f"{c}_方差" for c in selected])
    variances.insert(0, "EM类别", np.arange(1, final_k + 1))

    formulas = pd.DataFrame([
        {
            "项目": "EM-GMM模型",
            "说明": "样品来自K个高斯组分的混合分布，每个组分可解释为一个地球化学背景总体或混合背景单元。"
        },
        {
            "项目": "多随机种子试验",
            "说明": "同一K值下使用多个随机种子重复拟合GMM，并选取AIC最小的模型作为该K最佳模型。"
        },
        {
            "项目": "AIC",
            "说明": "AIC = 2p - 2ln(L)。AIC越小，表示模型拟合与复杂度折中越优。"
        },
        {
            "项目": "diag协方差自由参数数",
            "说明": "p = (K-1) + Kd + Kd = (K-1) + 2Kd，其中d为变量数。"
        },
        {
            "项目": "后验概率",
            "说明": "每个样品属于各EM类别的概率。最大后验概率越高，分类越稳定。"
        },
        {
            "项目": "分类不确定性",
            "说明": "分类不确定性 = 1 - 最大后验概率。数值越高，样品越可能位于过渡背景或混合背景。"
        },
        {
            "项目": "概率熵",
            "说明": "概率熵 = -sum(p_i ln p_i)。熵越高，说明样品类别归属越不确定。"
        },
    ])

    return {
        "模型参数": model_info_df,
        "混合权重": weights,
        "各类均值": means,
        "各类方差": variances,
        "计算公式说明": formulas,
    }


def build_paper_parameter_table(
    curve_df: pd.DataFrame,
    model_info_df: pd.DataFrame,
    selected: list[str],
    class_count_df: pd.DataFrame,
) -> pd.DataFrame:
    info_dict = dict(zip(model_info_df["参数"], model_info_df["取值"]))

    best_aic_k = int(curve_df.loc[curve_df["AIC"].idxmin(), "K"])
    best_aic_value = float(curve_df["AIC"].min())
    best_seed = int(curve_df.loc[curve_df["AIC"].idxmin(), "该K最佳随机种子"])

    min_class_n = int(class_count_df["样品数"].min())
    max_class_n = int(class_count_df["样品数"].max())

    class_count_text = "；".join(
        [
            f"类别{int(row['EM类别'])}: n={int(row['样品数'])}, {row['样品比例_%']:.2f}%"
            for _, row in class_count_df.iterrows()
        ]
    )

    return pd.DataFrame([{
        "输出K值": info_dict.get("最终K值"),
        "AIC曲线最小K": best_aic_k,
        "AIC曲线最小AIC": best_aic_value,
        "AIC曲线最小K对应最佳seed": best_seed,
        "该K模型AIC": info_dict.get("AIC"),
        "该K最佳随机种子": info_dict.get("该K最佳随机种子"),
        "随机种子试验数量": info_dict.get("随机种子试验数量"),
        "随机种子列表": info_dict.get("随机种子列表"),
        "参与元素数": len(selected),
        "样品数": info_dict.get("样品数"),
        "变量数": info_dict.get("变量数"),
        "类别数": info_dict.get("最终K值"),
        "最小类别样品数": min_class_n,
        "最大类别样品数": max_class_n,
        "各类别样品数": class_count_text,
        "协方差类型": info_dict.get("协方差类型"),
        "初始化方式": info_dict.get("初始化方式"),
        "最大迭代次数": info_dict.get("最大迭代次数"),
        "收敛容差": info_dict.get("收敛容差"),
        "正则化参数reg_covar": info_dict.get("正则化参数reg_covar"),
        "初始化次数n_init": info_dict.get("初始化次数n_init"),
        "是否收敛": info_dict.get("是否收敛"),
        "迭代次数": info_dict.get("迭代次数"),
        "log_likelihood_total": info_dict.get("log_likelihood_total"),
        "平均log_likelihood": info_dict.get("平均log_likelihood"),
        "理论自由参数数_diag": info_dict.get("理论自由参数数_diag"),
    }])


def build_report_key_table(
    paper_params: pd.DataFrame,
) -> pd.DataFrame:
    p = paper_params.iloc[0].to_dict()

    return pd.DataFrame([
        {"项目": "输出K值", "结果": p.get("输出K值")},
        {"项目": "AIC曲线最小K", "结果": p.get("AIC曲线最小K")},
        {"项目": "AIC曲线最小AIC", "结果": p.get("AIC曲线最小AIC")},
        {"项目": "AIC曲线最小K对应最佳seed", "结果": p.get("AIC曲线最小K对应最佳seed")},
        {"项目": "该K模型AIC", "结果": p.get("该K模型AIC")},
        {"项目": "该K最佳随机种子", "结果": p.get("该K最佳随机种子")},
        {"项目": "随机种子试验数量", "结果": p.get("随机种子试验数量")},
        {"项目": "有效样品数", "结果": p.get("样品数")},
        {"项目": "参与元素数", "结果": p.get("参与元素数")},
        {"项目": "最小类别样品数", "结果": p.get("最小类别样品数")},
        {"项目": "最大类别样品数", "结果": p.get("最大类别样品数")},
        {"项目": "是否收敛", "结果": p.get("是否收敛")},
        {"项目": "迭代次数", "结果": p.get("迭代次数")},
        {"项目": "各类别样品数", "结果": p.get("各类别样品数")},
    ])






def main():
    start_time = time.time()

    in_path = ask_file_path("请选择用于EM聚类的数据表")
    out_dir = ask_output_dir("EM聚类AIC多种子分析结果")

    df = read_table(in_path)
    df.columns = make_unique_columns(df.columns)

    coords, elements, coord_cols, element_cols = split_geochem_table(df)

    # Leave the coordinate columns out of the model fit.
    selected = element_cols

    X_df = elements[selected].apply(pd.to_numeric, errors="coerce")
    X_df = X_df.replace([np.inf, -np.inf], np.nan)

    valid_mask = X_df.notna().all(axis=1)
    X_df_valid = X_df.loc[valid_mask].reset_index(drop=True)
    valid_original = df.loc[valid_mask].reset_index(drop=True)

    invalid = df.loc[~valid_mask].copy()
    if len(invalid):
        invalid.insert(0, "剔除原因", "参与EM的元素存在缺失/非数值/无穷值")

    X = X_df_valid.to_numpy(dtype=float)

    if len(X) < 3:
        raise ValueError("有效样品数不足，无法进行EM聚类。")

    max_allowed_k = max(2, len(X) - 1)

    k_min, k_max = ask_k_range(2, min(10, max_allowed_k))
    k_max = min(k_max, max_allowed_k)

    seed_list = ask_seed_list()

    tag = now_tag()
    base = safe_name(Path(in_path).stem)
    prefix = out_dir / f"{base}_EM_AIC_多种子_{tag}"

    print("开始计算多随机种子AIC曲线...")
    print(f"结果保存文件夹：{out_dir}")
    print(f"随机种子数量：{len(seed_list)}")
    print(f"随机种子列表：{seed_list}")

    curve_df, all_seed_trials_df, best_models = build_aic_curve_table_multi_seed(
        X=X,
        k_min=k_min,
        k_max=k_max,
        seed_list=seed_list,
    )

    if curve_df.empty:
        raise ValueError("AIC曲线数据为空，请检查K值范围和样品数。")

    aic_png = plot_aic_only(curve_df, prefix)

    print(f"AIC曲线已生成：{aic_png}")

    open_file(aic_png)

    best_aic_k = int(curve_df.loc[curve_df["AIC"].idxmin(), "K"])

    final_ks = ask_final_k_values(
        max_k=max_allowed_k,
        default_text=str(best_aic_k),
    )

    seed_df = build_seed_record_df(seed_list)

    aic_curve_xlsx = prefix.with_name(prefix.name + "_AIC曲线与多种子数据.xlsx")
    write_excel(aic_curve_xlsx, {
        "00_随机种子记录": seed_df,
        "01_AIC曲线最佳结果": curve_df,
        "02_所有seed试验记录": all_seed_trials_df,
        "03_输入元素描述统计": descriptive_stats(X_df_valid),
        "04_未参与样品": invalid,
    })

    print(f"AIC曲线与多种子试验数据已保存：{aic_curve_xlsx}")

    generated_files = []

    for final_k in final_ks:
        print(f"开始生成K={final_k}的EM聚类分析结果...")

        if final_k in best_models:
            model = best_models[final_k]
        else:
            model, trial_df_extra = fit_best_gmm_by_seeds(X, final_k, seed_list)
            all_seed_trials_df = pd.concat([all_seed_trials_df, trial_df_extra], ignore_index=True)

        best_seed_for_k = int(
            all_seed_trials_df.loc[
                all_seed_trials_df["K"].eq(final_k)
                & all_seed_trials_df["是否该K最佳seed"].eq("是"),
                "随机种子"
            ].iloc[0]
        )

        print(f"K={final_k}最佳随机种子={best_seed_for_k}")

        result = build_cluster_result(
            valid_original=valid_original,
            coord_cols=coord_cols,
            model=model,
            X=X,
        )

        model_tables = build_model_tables(
            model=model,
            X=X,
            selected=selected,
            final_k=final_k,
            best_seed=best_seed_for_k,
            seed_list=seed_list,
            aic_png=aic_png,
        )

        weights_df = model_tables["混合权重"]
        class_count_df = build_concise_class_summary(result, weights_df=weights_df)

        stats_base = pd.concat(
            [
                result[
                    [
                        "EM类别",
                        "EM最大后验概率",
                        "EM分类不确定性",
                        "EM概率熵",
                    ]
                ].reset_index(drop=True),
                X_df_valid.reset_index(drop=True),
            ],
            axis=1,
        )

        detail_summary_df = cluster_summary_detail(stats_base, selected)

        means_df = model_tables["各类均值"]
        distance_df = class_mean_distance(means_df, selected)
        posterior_df = posterior_probability_summary(result, final_k)

        paper_params = build_paper_parameter_table(
            curve_df=curve_df,
            model_info_df=model_tables["模型参数"],
            selected=selected,
            class_count_df=class_count_df,
        )

        report_key_table = build_report_key_table(
            paper_params=paper_params,
        )

        this_k_seed_trials = all_seed_trials_df[
            all_seed_trials_df["K"].eq(final_k)
        ].copy().sort_values("AIC")

        out_xlsx = prefix.with_name(
            prefix.name + f"_K{final_k}_bestseed{best_seed_for_k}_分析结果.xlsx"
        )

        excel_sheets = {
            "00_随机种子记录": seed_df,
            "01_EM聚类结果": result,
            "02_类别样品数精简表": class_count_df,
            "03_报告核心参数": report_key_table,
            "04_AIC曲线最佳结果": curve_df,
            "05_该K所有seed试验": this_k_seed_trials,
            "06_所有seed试验记录": all_seed_trials_df,
            "07_模型参数": model_tables["模型参数"],
            "08_论文写作参数": paper_params,
            "09_混合权重": weights_df,
            "10_各类均值": model_tables["各类均值"],
            "11_各类方差": model_tables["各类方差"],
            "12_详细统计摘要": detail_summary_df,
            "13_后验概率统计": posterior_df,
            "14_类别均值距离": distance_df,
            "15_输入元素描述统计": descriptive_stats(X_df_valid),
            "16_未参与样品": invalid,
            "17_计算公式说明": model_tables["计算公式说明"],
        }

        write_excel(out_xlsx, excel_sheets)

        generated_files.append(out_xlsx)

        print(f"K={final_k}结果表：{out_xlsx}")

    elapsed = time.time() - start_time

    print("全部完成。")
    print(f"新生成结果文件夹：{out_dir}")
    print(f"AIC曲线：{aic_png}")
    print(f"AIC曲线与多种子数据：{aic_curve_xlsx}")
    print(f"运行耗时：{elapsed / 60:.2f} 分钟")

    for xlsx in generated_files:
        print(f"结果表：{xlsx}")

    messagebox.showinfo(
        "完成",
        (
            f"EM聚类多随机种子分析完成！\n\n"
            f"所有结果已保存至新生成文件夹：\n{out_dir}\n\n"
            f"AIC曲线已生成并打开：\n{aic_png}\n\n"
            f"随机种子数量：{len(seed_list)}\n"
            f"随机种子列表：{seed_list}\n\n"
            f"已按K分别生成Excel分析结果。\n\n"
            f"运行耗时：{elapsed / 60:.2f} 分钟"
        )
    )

    open_output_folder(out_dir)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("程序运行出错",f"{e}")
