# coding=utf-8

"""Perform MCD-based robust PCA after CLR and ILR transformations.

Input: Excel; two coordinate columns followed by positive element values.
Options: MCD alpha, random seed, loading cutoff, and output location.
Output: Detailed and summary Excel workbooks containing PC results.
Run: Start this script and follow the file and parameter dialogs."""

from __future__ import annotations

import math
import os
import re
import sys
import secrets
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk
except Exception as exc:
    raise RuntimeError(f"Tkinter 导入失败：{exc}") from exc

try:
    from sklearn.covariance import MinCovDet
except Exception as exc:
    raise RuntimeError(
        "scikit-learn 导入失败，请先安装：pip install scikit-learn\n"
        f"原始错误：{exc}"
    ) from exc

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except Exception as exc:
    raise RuntimeError(
        "openpyxl 导入失败，请先安装：pip install openpyxl\n"
        f"原始错误：{exc}"
    ) from exc


def 当前时间标签() -> str:
    """Return a timestamp for output names."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def 安全文件名(text: str) -> str:
    """Remove characters that are invalid in file names."""
    text = str(text).strip()
    if not text:
        text = "ILR_MCD_RPCA_PC1_PC20"

    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    text = re.sub(r"\s+", "_", text)

    return text[:120]


def 生成唯一列名(columns: list[Any]) -> list[str]:
    """Make duplicate column names unique."""
    seen: dict[str, int] = {}
    result: list[str] = []

    for col in columns:
        name = "" if col is None else str(col).strip()

        if not name:
            name = "未命名列"

        if name not in seen:
            seen[name] = 0
            result.append(name)
        else:
            seen[name] += 1
            result.append(f"{name}_重复{seen[name]}")

    return result


def 选择输入文件() -> Path:
    root = tk.Tk()
    root.withdraw()

    path = filedialog.askopenfilename(
        title="请选择 ILR + MCD-RPCA 输入Excel",
        filetypes=[
            ("Excel文件", "*.xlsx;*.xlsm"),
            ("所有文件", "*.*"),
        ],
    )

    root.destroy()

    if not path:
        raise SystemExit("已取消选择输入文件。")

    return Path(path)


def 选择工作表(path: Path) -> str:
    excel = pd.ExcelFile(path)
    sheets = excel.sheet_names

    if not sheets:
        raise ValueError("Excel中没有可用工作表。")

    if len(sheets) == 1:
        return sheets[0]

    root = tk.Tk()
    root.title("选择工作表")
    root.geometry("470x185")
    root.resizable(False, False)

    result = {"sheet": None}
    selected = tk.StringVar(value=sheets[0])

    tk.Label(
        root,
        text="请选择进行 ILR + MCD-RPCA 的工作表：",
        font=("Microsoft YaHei", 10),
    ).pack(padx=20, pady=(20, 8), anchor="w")

    combo = ttk.Combobox(
        root,
        textvariable=selected,
        values=sheets,
        state="readonly",
        width=52,
    )
    combo.pack(padx=20, pady=5)

    def 确定():
        result["sheet"] = selected.get()
        root.destroy()

    def 取消():
        root.destroy()

    frame = tk.Frame(root)
    frame.pack(pady=18)

    tk.Button(frame, text="确定", width=12, command=确定).pack(side="left", padx=8)
    tk.Button(frame, text="取消", width=12, command=取消).pack(side="left", padx=8)

    root.mainloop()

    if not result["sheet"]:
        raise SystemExit("已取消选择工作表。")

    return str(result["sheet"])


def 选择输出文件夹(default_dir: Path) -> Path:
    root = tk.Tk()
    root.withdraw()

    path = filedialog.askdirectory(
        title="请选择结果保存文件夹",
        initialdir=str(default_dir),
    )

    root.destroy()

    if not path:
        raise SystemExit("已取消选择输出文件夹。")

    return Path(path)


def 输入结果名称(default_name: str) -> str:
    root = tk.Tk()
    root.withdraw()

    name = simpledialog.askstring(
        "结果文件命名",
        "请输入本次结果名称：",
        initialvalue=default_name,
        parent=root,
    )

    root.destroy()

    if name is None:
        raise SystemExit("已取消结果命名。")

    return 安全文件名(name)


def 输入参数() -> dict[str, Any]:
    """Read the MCD and loading options."""
    root = tk.Tk()
    root.title("ILR + MCD-based Robust PCA 参数设置")
    root.geometry("790x465")
    root.resizable(False, False)

    alpha_var = tk.StringVar(value="0.50")
    seed_var = tk.StringVar(value="random")
    threshold_var = tk.StringVar(value="auto")

    row = 0

    tk.Label(
        root,
        text="ILR + MCD-based Robust PCA 参数设置",
        font=("Microsoft YaHei", 14, "bold"),
    ).grid(row=row, column=0, columnspan=3, pady=(18, 18))

    row += 1
    tk.Label(root, text="MCD alpha：", width=22, anchor="e").grid(
        row=row, column=0, padx=8, pady=9
    )
    tk.Entry(root, textvariable=alpha_var, width=24).grid(
        row=row, column=1, padx=8, pady=9, sticky="w"
    )
    tk.Label(
        root,
        text="范围0.50–1.00；推荐默认0.50",
        fg="#555555",
    ).grid(row=row, column=2, sticky="w")

    row += 1
    tk.Label(root, text="随机种子：", width=22, anchor="e").grid(
        row=row, column=0, padx=8, pady=9
    )
    tk.Entry(root, textvariable=seed_var, width=24).grid(
        row=row, column=1, padx=8, pady=9, sticky="w"
    )
    tk.Label(
        root,
        text="默认random：每次运行自动生成整数种子；也可手动输入整数",
        fg="#555555",
    ).grid(row=row, column=2, sticky="w")

    row += 1
    tk.Label(root, text="输出PC范围：", width=22, anchor="e").grid(
        row=row, column=0, padx=8, pady=9
    )
    tk.Label(
        root,
        text="PC1-PC20",
        font=("Microsoft YaHei", 10, "bold"),
    ).grid(row=row, column=1, padx=8, pady=9, sticky="w")
    tk.Label(
        root,
        text="若D-1<20，则自动输出全部可计算PC",
        fg="#555555",
    ).grid(row=row, column=2, sticky="w")

    row += 1
    tk.Label(root, text="候选载荷阈值：", width=22, anchor="e").grid(
        row=row, column=0, padx=8, pady=9
    )
    tk.Entry(root, textvariable=threshold_var, width=24).grid(
        row=row, column=1, padx=8, pady=9, sticky="w"
    )
    tk.Label(
        root,
        text="auto=1/sqrt(D)；0表示不筛候选，但全部元素仍输出",
        fg="#555555",
    ).grid(row=row, column=2, sticky="w")

    row += 1

    note = (
        "数据前提：已完成单位统一与缺失值/检出限处理。\n"
        "程序仍会检查所有元素是否严格 > 0；ILR不能处理0或负值。\n"
        "MCD的h由程序根据 n、D、alpha 自动计算，不需要手动填写。\n"
        "PC1-PC20全部元素载荷都会完整回映射到原始元素，不会只保留候选元素。"
    )

    tk.Label(
        root,
        text=note,
        justify="left",
        fg="#8B0000",
        anchor="w",
    ).grid(
        row=row,
        column=0,
        columnspan=3,
        padx=32,
        pady=(20, 10),
        sticky="w",
    )

    result: dict[str, Any] = {}

    def 确定():
        try:
            alpha = float(alpha_var.get().strip())

            if not 0.50 <= alpha <= 1.00:
                raise ValueError("MCD alpha必须位于0.50–1.00之间。")

            seed_text = seed_var.get().strip().lower()

            if seed_text in {"", "random", "rand", "随机"}:


                seed = secrets.randbelow(2_147_483_646) + 1
            else:
                seed = int(seed_text)

                if seed < 0:
                    raise ValueError("随机种子必须为非负整数，或输入 random / 随机。")

            threshold_text = threshold_var.get().strip().lower()

            if threshold_text == "auto":
                threshold: float | str = "auto"
            else:
                threshold = float(threshold_text)

                if threshold < 0:
                    raise ValueError("候选载荷阈值不能小于0。")

            result.update({
                "alpha": alpha,
                "seed": seed,
                "loading_threshold": threshold,
            })

            root.destroy()

        except Exception as exc:
            messagebox.showerror("参数错误", str(exc), parent=root)

    def 取消():
        root.destroy()

    frame = tk.Frame(root)
    frame.grid(row=row + 1, column=0, columnspan=3, pady=18)

    tk.Button(
        frame,
        text="开始计算",
        width=15,
        command=确定,
    ).pack(side="left", padx=10)

    tk.Button(
        frame,
        text="取消",
        width=12,
        command=取消,
    ).pack(side="left", padx=10)

    root.mainloop()

    if not result:
        raise SystemExit("已取消参数设置。")

    return result


def 读取输入数据(
    path: Path,
    sheet_name: str,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    list[str],
    list[str]
]:
    """Read element data and check for missing or nonpositive values."""
    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
    )

    if df.shape[1] < 4:
        raise ValueError(
            "至少需要4列：前2列坐标 + 至少2个元素变量。"
        )

    df.columns = 生成唯一列名(
        list(df.columns)
    )

    坐标列 = list(df.columns[:2])
    元素列 = list(df.columns[2:])

    坐标df = df.iloc[:, :2].copy()
    元素df = df.iloc[:, 2:].copy()


    元素数值df = 元素df.apply(
        pd.to_numeric,
        errors="coerce",
    )


    bad_mask = 元素数值df.isna()

    if bad_mask.any().any():
        positions = np.argwhere(
            bad_mask.to_numpy()
        )

        lines = []

        for i, j in positions[:20]:
            lines.append(
                f"Excel第{i+2}行，元素 {元素列[j]}：缺失值或非数值"
            )

        more = ""

        if len(positions) > 20:
            more = f"\n……另有 {len(positions)-20} 项"

        raise ValueError(
            "检测到缺失值或非数值。\n"
            "你已说明数据完成缺失值处理，请检查输入文件是否正确。\n\n"
            + "\n".join(lines)
            + more
        )

    X = 元素数值df.to_numpy(
        dtype=float
    )


    if not np.all(
        np.isfinite(X)
    ):
        raise ValueError(
            "元素数据中存在 NaN 或 Inf，不能进入ILR。"
        )


    if np.any(X <= 0):
        positions = np.argwhere(
            X <= 0
        )

        lines = []

        for i, j in positions[:20]:
            lines.append(
                f"Excel第{i+2}行，元素 {元素列[j]}，值={X[i, j]}"
            )

        more = ""

        if len(positions) > 20:
            more = f"\n……另有 {len(positions)-20} 项"

        raise ValueError(
            "ILR要求所有元素严格 > 0，但检测到0或负值。\n\n"
            + "\n".join(lines)
            + more
            + "\n\n请确认0值、低于检出限值和负值编码均已处理。"
        )

    return (
        df,
        坐标df,
        元素数值df,
        坐标列,
        元素列,
    )


def 构造Pivot_ILR正交基(
    D: int
) -> np.ndarray:
    """Build the orthonormal pivot ILR basis."""
    if D < 2:
        raise ValueError(
            "ILR至少需要2个元素。"
        )

    H = np.zeros(
        (D, D - 1),
        dtype=float,
    )

    for j in range(D - 1):
        remaining = D - j - 1

        H[j, j] = math.sqrt(
            remaining
            / (remaining + 1.0)
        )

        H[j + 1:, j] = (
            -1.0
            / math.sqrt(
                remaining
                * (remaining + 1.0)
            )
        )


    if not np.allclose(
        H.T @ H,
        np.eye(D - 1),
        atol=1e-10,
    ):
        raise RuntimeError(
            "ILR正交基构造失败：H.T @ H != I。"
        )


    if not np.allclose(
        H.sum(axis=0),
        0.0,
        atol=1e-10,
    ):
        raise RuntimeError(
            "ILR正交基构造失败：基向量不满足CLR零和约束。"
        )

    return H


def CLR变换(
    X: np.ndarray
) -> np.ndarray:
    """Compute centered log-ratio coordinates."""
    logX = np.log(X)

    return (
        logX
        - np.mean(
            logX,
            axis=1,
            keepdims=True,
        )
    )


def ILR变换(
    X: np.ndarray,
    H: np.ndarray,
) -> tuple[
    np.ndarray,
    np.ndarray,
    float
]:
    """Transform CLR coordinates into ILR and check reconstruction."""
    CLR = CLR变换(X)

    ILR = CLR @ H


    CLR_reconstructed = (
        ILR @ H.T
    )

    max_error = float(
        np.max(
            np.abs(
                CLR
                - CLR_reconstructed
            )
        )
    )

    return (
        ILR,
        CLR,
        max_error,
    )


def 自动计算MCD_h(
    alpha: float,
    n: int,
    p: int,
) -> int:
    """Determine the MCD support size from n, p, and alpha."""
    if not 0.50 <= alpha <= 1.00:
        raise ValueError(
            "alpha必须位于0.50–1.00。"
        )

    n2 = (
        n + p + 1
    ) // 2

    h = math.floor(
        2 * n2
        - n
        + 2.0
        * (n - n2)
        * alpha
    )


    h = max(
        p + 1,
        min(
            h,
            n,
        )
    )

    return int(h)


def 运行ILR_MCD_RPCA(
    X: np.ndarray,
    元素列: list[str],
    alpha: float,
    seed: int,
) -> dict[str, Any]:
    """Estimate MCD-based PCs in ILR space and map loadings to CLR space."""
    n, D = X.shape

    p = D - 1


    if n <= p:
        raise ValueError(
            f"样品数不足：n={n}，ILR维数p={p}。"
            "MCD要求样品数大于ILR维数。"
        )


    维度比警告 = ""

    if n / p < 2.0:
        维度比警告 = (
            f"当前 n/p={n/p:.2f}，样品数相对于ILR维度偏少，"
            "MCD协方差及后部PC可能较不稳定，建议谨慎解释。"
        )


    H = 构造Pivot_ILR正交基(D)

    ILR, CLR, clr_reconstruction_error = ILR变换(
        X=X,
        H=H,
    )


    h = 自动计算MCD_h(
        alpha=alpha,
        n=n,
        p=p,
    )

    support_fraction = (
        h / n
    )


    with warnings.catch_warnings(
        record=True
    ) as caught:

        warnings.simplefilter(
            "always"
        )

        mcd = MinCovDet(
            store_precision=True,
            assume_centered=False,
            support_fraction=support_fraction,
            random_state=seed,
        )

        mcd.fit(
            ILR
        )

    warning_messages = [
        str(w.message)
        for w in caught
    ]

    if 维度比警告:
        warning_messages.insert(
            0,
            维度比警告,
        )

    robust_location = np.asarray(
        mcd.location_,
        dtype=float,
    )

    robust_cov = np.asarray(
        mcd.covariance_,
        dtype=float,
    )

    raw_location = np.asarray(
        mcd.raw_location_,
        dtype=float,
    )

    raw_cov = np.asarray(
        mcd.raw_covariance_,
        dtype=float,
    )

    raw_support = np.asarray(
        mcd.raw_support_,
        dtype=bool,
    )

    final_support = np.asarray(
        mcd.support_,
        dtype=bool,
    )

    robust_distance_sq = np.asarray(
        mcd.dist_,
        dtype=float,
    )

    robust_distance = np.sqrt(
        np.maximum(
            robust_distance_sq,
            0.0,
        )
    )


    eigenvalues, eigenvectors = np.linalg.eigh(
        robust_cov
    )


    order = np.argsort(
        eigenvalues
    )[::-1]

    eigenvalues = eigenvalues[
        order
    ]

    eigenvectors = eigenvectors[
        :,
        order
    ]


    max_eval = (
        float(
            np.max(
                np.abs(
                    eigenvalues
                )
            )
        )
        if eigenvalues.size
        else 0.0
    )

    tol = (
        max_eval
        * max(1, p)
        * np.finfo(float).eps
        * 100.0
    )

    tiny_negative = (
        (eigenvalues < 0)
        &
        (np.abs(eigenvalues) <= tol)
    )

    eigenvalues[
        tiny_negative
    ] = 0.0

    if np.any(
        eigenvalues < -tol
    ):
        raise ValueError(
            "MCD稳健协方差出现明显负特征值。"
            "请检查变量共线性、样品数和输入数据。"
        )

    total_variance = float(
        np.sum(
            eigenvalues
        )
    )

    if total_variance <= 0:
        raise ValueError(
            "MCD稳健协方差总方差<=0，无法进行PCA。"
        )

    explained_ratio = (
        eigenvalues
        / total_variance
    )

    cumulative_ratio = np.cumsum(
        explained_ratio
    )


    centered_ILR = (
        ILR
        - robust_location
    )

    scores_all = (
        centered_ILR
        @ eigenvectors
    )


    ilr_loadings_all = (
        eigenvectors.copy()
    )


    # Map the PC directions from ILR to CLR coordinates.
    clr_loadings_all = (
        H
        @ ilr_loadings_all
    )


    robust_center_clr = (
        robust_location
        @ H.T
    )


    centered_CLR = (
        centered_ILR
        @ H.T
    )


    if clr_loadings_all.shape[0] != D:
        raise RuntimeError(
            "CLR元素载荷回映射失败："
            f"输入D={D}个元素，回映射得到"
            f"{clr_loadings_all.shape[0]}个元素。"
        )

    if centered_CLR.shape[1] != D:
        raise RuntimeError(
            "CLR中心化数据回映射失败："
            f"应有D={D}列，实际为{centered_CLR.shape[1]}列。"
        )

    if robust_center_clr.shape[0] != D:
        raise RuntimeError(
            "CLR稳健中心回映射失败："
            f"应有D={D}个元素，实际为{robust_center_clr.shape[0]}个。"
        )


    clr_loading_sums_before_sign = np.sum(
        clr_loadings_all,
        axis=0,
    )

    if not np.allclose(
        clr_loading_sums_before_sign,
        0.0,
        atol=1e-8,
    ):
        raise RuntimeError(
            "CLR载荷回映射后未满足零和约束。"
        )


    # Fix the sign convention for consistent PC interpretation.
    sign_flips: list[int] = []

    for k in range(
        clr_loadings_all.shape[1]
    ):
        idx = int(
            np.argmax(
                np.abs(
                    clr_loadings_all[
                        :,
                        k
                    ]
                )
            )
        )

        if (
            clr_loadings_all[
                idx,
                k
            ] < 0
        ):
            clr_loadings_all[
                :,
                k
            ] *= -1.0

            ilr_loadings_all[
                :,
                k
            ] *= -1.0

            scores_all[
                :,
                k
            ] *= -1.0

            sign_flips.append(
                -1
            )
        else:
            sign_flips.append(
                1
            )


    clr_loading_sums = np.sum(
        clr_loadings_all,
        axis=0,
    )

    if not np.allclose(
        clr_loading_sums,
        0.0,
        atol=1e-8,
    ):
        raise RuntimeError(
            "PC符号调整后CLR载荷零和约束检查失败。"
        )


    q = min(
        20,
        p,
    )

    return {
        "H": H,
        "ILR": ILR,
        "CLR": CLR,
        "centered_ILR": centered_ILR,
        "centered_CLR": centered_CLR,

        "h": h,
        "support_fraction": support_fraction,

        "robust_location": robust_location,
        "robust_center_clr": robust_center_clr,
        "raw_location": raw_location,

        "robust_cov": robust_cov,
        "raw_cov": raw_cov,

        "raw_support": raw_support,
        "final_support": final_support,

        "robust_distance_sq": robust_distance_sq,
        "robust_distance": robust_distance,

        "eigenvalues": eigenvalues,
        "explained_ratio": explained_ratio,
        "cumulative_ratio": cumulative_ratio,

        "scores_all": scores_all,
        "ilr_loadings_all": ilr_loadings_all,
        "clr_loadings_all": clr_loadings_all,

        "sign_flips": sign_flips,
        "warning_messages": warning_messages,

        "q": q,

        "clr_loading_sums": clr_loading_sums,
        "clr_reconstruction_error": clr_reconstruction_error,
        "元素列": 元素列,
    }


def 构建元素PC综合表(
    元素列: list[str],
    clr_loadings: np.ndarray,
    explained_ratio: np.ndarray,
    q: int,
    threshold: float,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame
]:
    """Summarize element loadings and selected candidates."""
    D = len(
        元素列
    )

    if (
        clr_loadings.shape[0]
        != D
    ):
        raise RuntimeError(
            "CLR元素载荷元素数与原始元素数不一致。"
        )

    综合rows = []
    最大PCrows = []
    候选rows = []

    for i, element in enumerate(
        元素列
    ):
        row: dict[
            str,
            Any
        ] = {
            "元素": element
        }

        abs_values = []

        for k in range(q):
            loading = float(
                clr_loadings[
                    i,
                    k
                ]
            )

            abs_loading = abs(
                loading
            )

            contribution = (
                loading ** 2
            )

            row[
                f"PC{k+1}_CLR载荷"
            ] = loading

            row[
                f"PC{k+1}_绝对载荷"
            ] = abs_loading

            row[
                f"PC{k+1}_平方贡献"
            ] = contribution

            row[
                f"PC{k+1}_方向"
            ] = (
                "正向"
                if loading > 0
                else (
                    "负向"
                    if loading < 0
                    else "零"
                )
            )

            row[
                f"PC{k+1}_达到候选阈值"
            ] = (
                "是"
                if (
                    threshold > 0
                    and
                    abs_loading >= threshold
                )
                else "否"
            )

            abs_values.append(
                abs_loading
            )

            if (
                threshold > 0
                and
                abs_loading >= threshold
            ):
                候选rows.append({
                    "主成分": f"PC{k+1}",
                    "PC解释方差百分比(%)": explained_ratio[k] * 100,
                    "元素": element,
                    "CLR载荷": loading,
                    "绝对载荷": abs_loading,
                    "平方贡献": contribution,
                    "方向": (
                        "正向"
                        if loading > 0
                        else "负向"
                    ),
                    "候选阈值": threshold,
                })

        dominant_idx = int(
            np.argmax(
                abs_values
            )
        )

        row[
            "PC1-PC20中最大载荷所在PC"
        ] = f"PC{dominant_idx+1}"

        row[
            "PC1-PC20中最大绝对载荷"
        ] = float(
            abs_values[
                dominant_idx
            ]
        )

        row[
            "最大载荷方向"
        ] = (
            "正向"
            if clr_loadings[
                i,
                dominant_idx
            ] > 0
            else "负向"
        )

        综合rows.append(
            row
        )

        最大PCrows.append({
            "元素": element,
            "最大绝对载荷所在PC": f"PC{dominant_idx+1}",
            "该PC解释方差百分比(%)": explained_ratio[dominant_idx] * 100,
            "CLR载荷": float(
                clr_loadings[
                    i,
                    dominant_idx
                ]
            ),
            "绝对载荷": float(
                abs_values[
                    dominant_idx
                ]
            ),
            "方向": (
                "正向"
                if clr_loadings[
                    i,
                    dominant_idx
                ] > 0
                else "负向"
            ),
        })

    综合df = pd.DataFrame(
        综合rows
    )

    最大PCdf = pd.DataFrame(
        最大PCrows
    ).sort_values(
        "绝对载荷",
        ascending=False,
    ).reset_index(
        drop=True
    )

    候选df = pd.DataFrame(
        候选rows
    )

    if not 候选df.empty:
        候选df[
            "PC序号"
        ] = (
            候选df[
                "主成分"
            ]
            .str.replace(
                "PC",
                "",
                regex=False,
            )
            .astype(int)
        )

        候选df = (
            候选df
            .sort_values(
                [
                    "PC序号",
                    "绝对载荷",
                ],
                ascending=[
                    True,
                    False,
                ],
            )
            .drop(
                columns=[
                    "PC序号"
                ]
            )
            .reset_index(
                drop=True
            )
        )

    return (
        综合df,
        最大PCdf,
        候选df,
    )


def 构建全部元素载荷长表(
    元素列: list[str],
    clr_loadings: np.ndarray,
    explained_ratio: np.ndarray,
    q: int,
    threshold: float,
) -> pd.DataFrame:
    """List the loadings of every element for each reported PC."""
    D = len(元素列)

    if clr_loadings.shape[0] != D:
        raise RuntimeError(
            "CLR元素载荷元素数与原始元素数不一致，"
            "无法生成全部元素载荷长表。"
        )

    rows = []

    for k in range(q):
        for i, element in enumerate(元素列):
            loading = float(
                clr_loadings[
                    i,
                    k
                ]
            )

            abs_loading = abs(
                loading
            )

            rows.append({
                "主成分": f"PC{k+1}",
                "PC解释方差百分比(%)": explained_ratio[k] * 100,
                "元素": element,
                "CLR载荷": loading,
                "绝对载荷": abs_loading,
                "平方贡献": loading ** 2,
                "方向": (
                    "正向"
                    if loading > 0
                    else (
                        "负向"
                        if loading < 0
                        else "零"
                    )
                ),
                "候选阈值": threshold,
            })

    df = pd.DataFrame(
        rows
    )

    if not df.empty:
        df[
            "PC序号"
        ] = (
            df[
                "主成分"
            ]
            .str.replace(
                "PC",
                "",
                regex=False,
            )
            .astype(int)
        )

        df = (
            df
            .sort_values(
                [
                    "PC序号",
                    "绝对载荷",
                ],
                ascending=[
                    True,
                    False,
                ],
            )
            .drop(
                columns=[
                    "PC序号"
                ]
            )
            .reset_index(
                drop=True
            )
        )

    return df


def 构建单PC全部元素表(
    元素列: list[str],
    clr_loadings: np.ndarray,
    explained_ratio: np.ndarray,
    pc_index: int,
    threshold: float,
) -> pd.DataFrame:
    """Sort all element loadings for one principal component."""
    k = pc_index

    df = pd.DataFrame({
        "元素": 元素列,
        "CLR载荷": clr_loadings[:, k],
    })

    df[
        "绝对载荷"
    ] = df[
        "CLR载荷"
    ].abs()

    df[
        "平方贡献"
    ] = (
        df[
            "CLR载荷"
        ] ** 2
    )

    df[
        "方向"
    ] = np.where(
        df[
            "CLR载荷"
        ] > 0,
        "正向",
        np.where(
            df[
                "CLR载荷"
            ] < 0,
            "负向",
            "零",
        )
    )

    df[
        "是否达到候选阈值"
    ] = np.where(
        (
            threshold > 0
        )
        &
        (
            df[
                "绝对载荷"
            ] >= threshold
        ),
        "是",
        "否",
    )

    df[
        "该PC解释方差百分比(%)"
    ] = (
        explained_ratio[k]
        * 100
    )


    df = (
        df
        .sort_values(
            "绝对载荷",
            ascending=False,
        )
        .reset_index(
            drop=True
        )
    )

    df.insert(
        0,
        "绝对载荷排序",
        np.arange(
            1,
            len(df) + 1
        )
    )

    return df


def 美化Excel工作簿(
    output_path: Path,
):
    """Apply basic formatting to an Excel workbook."""
    wb = load_workbook(
        output_path
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        if ws.max_row >= 1:
            ws.auto_filter.ref = (
                ws.dimensions
            )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for col_idx in range(
            1,
            ws.max_column + 1
        ):
            max_len = 10

            for row_idx in range(
                1,
                min(
                    ws.max_row,
                    200,
                ) + 1
            ):
                value = ws.cell(
                    row=row_idx,
                    column=col_idx,
                ).value

                if value is not None:
                    max_len = max(
                        max_len,
                        min(
                            len(
                                str(
                                    value
                                )
                            ) + 2,
                            42,
                        ),
                    )

            ws.column_dimensions[
                get_column_letter(
                    col_idx
                )
            ].width = (
                max_len
            )

    wb.save(
        output_path
    )


def 写入核心结果Excel(
    output_path: Path,
    元素列: list[str],
    result: dict[str, Any],
    threshold: float,
):
    """Write the three main result tables to Excel."""
    q = result[
        "q"
    ]


    PC方差rows = []

    for k in range(q):
        PC方差rows.append({
            "主成分": f"PC{k+1}",
            "特征值": result["eigenvalues"][k],
            "解释方差比例": result["explained_ratio"][k],
            "解释方差百分比(%)": result["explained_ratio"][k] * 100,
            "累计解释方差比例": result["cumulative_ratio"][k],
            "累计解释方差百分比(%)": result["cumulative_ratio"][k] * 100,
            "PC整体符号乘子": result["sign_flips"][k],
        })

    PC方差df = pd.DataFrame(
        PC方差rows
    )


    全部元素载荷df = 构建全部元素载荷长表(
        元素列=元素列,
        clr_loadings=result[
            "clr_loadings_all"
        ],
        explained_ratio=result[
            "explained_ratio"
        ],
        q=q,
        threshold=threshold,
    )


    (
        _元素综合df,
        _最大PCdf,
        候选df,
    ) = 构建元素PC综合表(
        元素列=元素列,
        clr_loadings=result[
            "clr_loadings_all"
        ],
        explained_ratio=result[
            "explained_ratio"
        ],
        q=q,
        threshold=threshold,
    )

    if 候选df.empty:


        候选df = pd.DataFrame(
            columns=[
                "主成分",
                "PC解释方差百分比(%)",
                "元素",
                "CLR载荷",
                "绝对载荷",
                "平方贡献",
                "方向",
                "候选阈值",
            ]
        )


    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:

        PC方差df.to_excel(
            writer,
            sheet_name="PC解释方差",
            index=False,
        )

        全部元素载荷df.to_excel(
            writer,
            sheet_name="元素载荷",
            index=False,
        )

        候选df.to_excel(
            writer,
            sheet_name="候选元素载荷",
            index=False,
        )

    美化Excel工作簿(
        output_path
    )


def 写入Excel结果(
    output_path: Path,
    input_path: Path,
    sheet_name: str,
    坐标df: pd.DataFrame,
    元素df: pd.DataFrame,
    坐标列: list[str],
    元素列: list[str],
    params: dict[str, Any],
    result: dict[str, Any],
    threshold: float,
):
    n = len(
        元素df
    )

    D = len(
        元素列
    )

    p = (
        D - 1
    )

    q = result[
        "q"
    ]


    参数df = pd.DataFrame(
        [
            ["输入文件", str(input_path), "输入Excel文件"],
            ["输入工作表", sheet_name, "本次读取Sheet"],
            ["坐标列", " / ".join(坐标列), "前2列，仅保留，不参与ILR/MCD/RPCA"],
            ["元素列", "第3列以后全部变量", f"共{D}个元素/氧化物变量"],
            ["输入元素数量D", D, "CLR回映射后必须仍为D个元素"],
            ["ILR维数p", p, "D-1"],
            ["输出主成分范围", f"PC1-PC{q}", "默认最多输出PC1-PC20"],
            ["ILR正交基", "Pivot sequential orthonormal basis", "D×(D-1)"],
            ["MCD实现", "sklearn.covariance.MinCovDet / FastMCD", "Python实现"],
            ["MCD alpha", params["alpha"], "用户设置；默认0.50"],
            ["MCD h", result["h"], "程序自动计算"],
            ["MCD support_fraction", result["support_fraction"], "h/n"],
            ["MCD随机种子", params["seed"], "默认random时由程序本次运行随机生成；记录该整数可用于结果复现"],
            ["MCD原始支持样本数", int(np.sum(result["raw_support"])), "raw_support_"],
            ["MCD最终支持样本数", int(np.sum(result["final_support"])), "support_"],
            ["PCA求解方式", "MCD最终稳健协方差矩阵特征值分解", "不是对原始元素做普通PCA"],
            ["ILR载荷回CLR公式", "CLR载荷 = H @ ILR载荷", "完整恢复全部原始元素"],
            ["CLR回映射元素数量", result["clr_loadings_all"].shape[0], "必须等于输入D"],
            ["候选元素载荷阈值", threshold, "仅辅助标记，不删除任何元素"],
            ["程序内部Z-score", "否", "ILR前不做Z-score"],
            ["输入正值要求", "是", "所有元素值必须严格>0"],
            ["CLR→ILR→CLR最大重构误差", result["clr_reconstruction_error"], "越接近0越好"],
        ],
        columns=[
            "参数",
            "数值",
            "说明",
        ],
    )


    输入检查rows = []

    for elem in 元素列:
        col = 元素df[
            elem
        ].to_numpy(
            dtype=float
        )

        输入检查rows.append({
            "元素": elem,
            "样品数": len(col),
            "最小值": float(np.min(col)),
            "最大值": float(np.max(col)),
            "均值": float(np.mean(col)),
            "样本标准差": float(np.std(col, ddof=1)) if len(col) > 1 else np.nan,
            "非正值数量": int(np.sum(col <= 0)),
            "非有限值数量": int(np.sum(~np.isfinite(col))),
        })

    输入检查df = pd.DataFrame(
        输入检查rows
    )


    H = result[
        "H"
    ]

    ILR基df = pd.DataFrame(
        H,
        index=元素列,
        columns=[
            f"ILR{i+1}"
            for i in range(p)
        ],
    ).reset_index().rename(
        columns={
            "index": "原始元素"
        }
    )


    ILR数据df = pd.concat(
        [
            坐标df.reset_index(
                drop=True
            ),
            pd.DataFrame(
                result["ILR"],
                columns=[
                    f"ILR{i+1}"
                    for i in range(p)
                ],
            ),
        ],
        axis=1,
    )


    CLR原始df = pd.concat(
        [
            坐标df.reset_index(
                drop=True
            ),
            pd.DataFrame(
                result["CLR"],
                columns=元素列,
            ),
        ],
        axis=1,
    )


    CLR中心化df = pd.concat(
        [
            坐标df.reset_index(
                drop=True
            ),
            pd.DataFrame(
                result[
                    "centered_CLR"
                ],
                columns=元素列,
            ),
        ],
        axis=1,
    )


    MCD汇总df = pd.DataFrame(
        [
            ["样品数n", n],
            ["元素数D", D],
            ["ILR维数p", p],
            ["alpha", params["alpha"]],
            ["自动计算h", result["h"]],
            ["support_fraction=h/n", result["support_fraction"]],
            ["原始MCD支持样本数", int(np.sum(result["raw_support"]))],
            ["最终MCD支持样本数", int(np.sum(result["final_support"]))],
            ["样品数/ILR维数(n/p)", n / p],
            [
                "运行警告",
                " | ".join(result["warning_messages"])
                if result["warning_messages"]
                else "无",
            ],
        ],
        columns=[
            "指标",
            "数值",
        ],
    )


    MCD中心df = pd.DataFrame({
        "ILR坐标": [
            f"ILR{i+1}"
            for i in range(p)
        ],
        "最终稳健中心": result[
            "robust_location"
        ],
        "原始MCD中心": result[
            "raw_location"
        ],
    })


    ilr_names = [
        f"ILR{i+1}"
        for i in range(p)
    ]

    MCD稳健协方差df = pd.DataFrame(
        result[
            "robust_cov"
        ],
        index=ilr_names,
        columns=ilr_names,
    ).reset_index().rename(
        columns={
            "index": "ILR坐标"
        }
    )

    MCD原始协方差df = pd.DataFrame(
        result[
            "raw_cov"
        ],
        index=ilr_names,
        columns=ilr_names,
    ).reset_index().rename(
        columns={
            "index": "ILR坐标"
        }
    )


    MCD支持df = 坐标df.reset_index(
        drop=True
    ).copy()

    MCD支持df[
        "原始MCD支持集"
    ] = result[
        "raw_support"
    ].astype(
        int
    )

    MCD支持df[
        "最终MCD支持集"
    ] = result[
        "final_support"
    ].astype(
        int
    )


    MCD距离df = 坐标df.reset_index(
        drop=True
    ).copy()

    MCD距离df[
        "MCD马氏距离"
    ] = result[
        "robust_distance"
    ]

    MCD距离df[
        "MCD马氏距离平方"
    ] = result[
        "robust_distance_sq"
    ]

    MCD距离df[
        "最终MCD支持集"
    ] = result[
        "final_support"
    ].astype(
        int
    )


    PC方差rows = []

    for k in range(q):
        PC方差rows.append({
            "主成分": f"PC{k+1}",
            "特征值": result["eigenvalues"][k],
            "解释方差比例": result["explained_ratio"][k],
            "解释方差百分比(%)": result["explained_ratio"][k] * 100,
            "累计解释方差比例": result["cumulative_ratio"][k],
            "累计解释方差百分比(%)": result["cumulative_ratio"][k] * 100,
            "PC整体符号乘子": result["sign_flips"][k],
        })

    PC方差df = pd.DataFrame(
        PC方差rows
    )


    PC得分df = pd.concat(
        [
            坐标df.reset_index(
                drop=True
            ),
            pd.DataFrame(
                result[
                    "scores_all"
                ][
                    :,
                    :q
                ],
                columns=[
                    f"PC{k+1}"
                    for k in range(q)
                ],
            ),
        ],
        axis=1,
    )


    ILR载荷df = pd.DataFrame(
        result[
            "ilr_loadings_all"
        ][
            :,
            :q
        ],
        index=ilr_names,
        columns=[
            f"PC{k+1}"
            for k in range(q)
        ],
    ).reset_index().rename(
        columns={
            "index": "ILR坐标"
        }
    )


    CLR载荷df = pd.DataFrame(
        result[
            "clr_loadings_all"
        ][
            :,
            :q
        ],
        index=元素列,
        columns=[
            f"PC{k+1}"
            for k in range(q)
        ],
    ).reset_index().rename(
        columns={
            "index": "元素"
        }
    )

    recovered_elements = (
        CLR载荷df[
            "元素"
        ].tolist()
    )

    if (
        recovered_elements
        != 元素列
    ):
        raise RuntimeError(
            "CLR元素载荷表中的元素顺序与原始输入不一致，停止输出。"
        )

    if (
        len(
            CLR载荷df
        )
        != D
    ):
        raise RuntimeError(
            f"CLR元素载荷应有{D}个元素，"
            f"实际为{len(CLR载荷df)}。"
        )


    (
        元素综合df,
        最大PCdf,
        候选df,
    ) = 构建元素PC综合表(
        元素列=元素列,
        clr_loadings=result["clr_loadings_all"],
        explained_ratio=result["explained_ratio"],
        q=q,
        threshold=threshold,
    )

    if 候选df.empty:
        候选df = pd.DataFrame({
            "说明": [
                "当前没有元素达到候选阈值，或阈值设置为0。"
                "全部元素载荷仍完整保存在“16_CLR元素载荷_全部”及PC01-PC20单独工作表。"
            ]
        })


    校验rows = [
        ["输入元素数量D", D, "必须与回映射元素数一致"],
        [
            "CLR回映射元素数量",
            result["clr_loadings_all"].shape[0],
            "必须等于D",
        ],
        [
            "输入元素列表",
            " | ".join(元素列),
            "原始Excel第3列以后",
        ],
        [
            "回映射元素列表",
            " | ".join(recovered_elements),
            "必须与输入完全一致",
        ],
        [
            "缺失元素数量",
            len(
                set(元素列)
                - set(recovered_elements)
            ),
            "必须为0",
        ],
        [
            "额外元素数量",
            len(
                set(recovered_elements)
                - set(元素列)
            ),
            "必须为0",
        ],
        [
            "元素顺序完全一致",
            (
                "是"
                if recovered_elements == 元素列
                else "否"
            ),
            "必须为是",
        ],
        [
            "CLR载荷矩阵形状",
            str(
                result[
                    "clr_loadings_all"
                ][
                    :,
                    :q
                ].shape
            ),
            f"应为({D}, {q})",
        ],
        [
            "CLR→ILR→CLR最大重构误差",
            result[
                "clr_reconstruction_error"
            ],
            "应非常接近0",
        ],
    ]

    for k in range(q):
        校验rows.append([
            f"PC{k+1}全部元素CLR载荷之和",
            float(
                np.sum(
                    result[
                        "clr_loadings_all"
                    ][
                        :,
                        k
                    ]
                )
            ),
            "CLR零和约束：应接近0",
        ])

    校验df = pd.DataFrame(
        校验rows,
        columns=[
            "校验项目",
            "结果",
            "判定要求/说明",
        ],
    )


    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:

        参数df.to_excel(
            writer,
            sheet_name="01_运行参数",
            index=False,
        )

        输入检查df.to_excel(
            writer,
            sheet_name="02_输入数据检查",
            index=False,
        )

        ILR基df.to_excel(
            writer,
            sheet_name="03_ILR正交基",
            index=False,
        )

        ILR数据df.to_excel(
            writer,
            sheet_name="04_ILR数据",
            index=False,
        )

        CLR原始df.to_excel(
            writer,
            sheet_name="05_CLR原始坐标",
            index=False,
        )

        CLR中心化df.to_excel(
            writer,
            sheet_name="06_CLR稳健中心化",
            index=False,
        )

        MCD汇总df.to_excel(
            writer,
            sheet_name="07_MCD汇总",
            index=False,
        )

        MCD中心df.to_excel(
            writer,
            sheet_name="08_MCD稳健中心",
            index=False,
        )

        MCD稳健协方差df.to_excel(
            writer,
            sheet_name="09_MCD稳健协方差",
            index=False,
        )

        MCD原始协方差df.to_excel(
            writer,
            sheet_name="10_MCD原始协方差",
            index=False,
        )

        MCD支持df.to_excel(
            writer,
            sheet_name="11_MCD支持集",
            index=False,
        )

        MCD距离df.to_excel(
            writer,
            sheet_name="12_MCD稳健距离",
            index=False,
        )

        PC方差df.to_excel(
            writer,
            sheet_name="13_PC解释方差",
            index=False,
        )

        PC得分df.to_excel(
            writer,
            sheet_name="14_样品PC得分",
            index=False,
        )

        ILR载荷df.to_excel(
            writer,
            sheet_name="15_ILR载荷",
            index=False,
        )

        CLR载荷df.to_excel(
            writer,
            sheet_name="16_CLR元素载荷_全部",
            index=False,
        )

        元素综合df.to_excel(
            writer,
            sheet_name="17_元素PC综合表",
            index=False,
        )

        最大PCdf.to_excel(
            writer,
            sheet_name="18_元素最大载荷PC",
            index=False,
        )

        候选df.to_excel(
            writer,
            sheet_name="19_候选元素载荷",
            index=False,
        )

        校验df.to_excel(
            writer,
            sheet_name="20_CLR回映射校验",
            index=False,
        )


        for k in range(q):
            单PCdf = 构建单PC全部元素表(
                元素列=元素列,
                clr_loadings=result["clr_loadings_all"],
                explained_ratio=result["explained_ratio"],
                pc_index=k,
                threshold=threshold,
            )

            单PCdf.to_excel(
                writer,
                sheet_name=f"PC{k+1:02d}_全部元素载荷",
                index=False,
            )


    美化Excel工作簿(
        output_path
    )


def main():
    try:


        input_path = 选择输入文件()

        if (
            input_path.suffix.lower()
            not in {
                ".xlsx",
                ".xlsm",
            }
        ):
            raise ValueError(
                "当前版本仅支持 .xlsx / .xlsm。"
            )


        sheet_name = 选择工作表(
            input_path
        )


        output_dir = 选择输出文件夹(
            input_path.parent
        )


        result_name = 输入结果名称(
            f"{input_path.stem}_ILR_MCD_RPCA_PC1_PC20"
        )


        params = 输入参数()


        (
            原始df,
            坐标df,
            元素df,
            坐标列,
            元素列,
        ) = 读取输入数据(
            input_path,
            sheet_name,
        )

        X = 元素df.to_numpy(
            dtype=float
        )

        n, D = X.shape

        p = (
            D - 1
        )

        if D < 2:
            raise ValueError(
                "至少需要2个元素进行ILR。"
            )

        if n <= p:
            raise ValueError(
                f"当前样品数n={n}，ILR维数p={p}。"
                "样品数必须大于ILR维数。"
            )


        if (
            params[
                "loading_threshold"
            ]
            == "auto"
        ):
            threshold = (
                1.0
                / math.sqrt(
                    D
                )
            )
        else:
            threshold = float(
                params[
                    "loading_threshold"
                ]
            )


        result = 运行ILR_MCD_RPCA(
            X=X,
            元素列=元素列,
            alpha=params["alpha"],
            seed=params["seed"],
        )

        q = result[
            "q"
        ]



        output_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        tag = 当前时间标签()

        excel_path = (
            output_dir
            / f"{result_name}_{tag}.xlsx"
        )

        核心excel_path = (
            output_dir
            / f"{result_name}_{tag}_核心结果.xlsx"
        )


        写入Excel结果(
            output_path=excel_path,
            input_path=input_path,
            sheet_name=sheet_name,
            坐标df=坐标df,
            元素df=元素df,
            坐标列=坐标列,
            元素列=元素列,
            params=params,
            result=result,
            threshold=threshold,
        )


        写入核心结果Excel(
            output_path=核心excel_path,
            元素列=元素列,
            result=result,
            threshold=threshold,
        )


        print(
            "\n"
            + "=" * 80
        )

        print(
            "ILR + MCD-based Robust PCA 计算完成"
        )

        print(
            "=" * 80
        )

        print(
            f"样品数 n：{n}"
        )

        print(
            f"输入元素数 D：{D}"
        )

        print(
            f"ILR维数 p：{p}"
        )

        print(
            f"输出PC范围：PC1-PC{q}"
        )

        print(
            f"MCD alpha：{params['alpha']}"
        )

        print(
            f"MCD本次实际随机种子：{params['seed']}"
        )

        print(
            f"自动计算 h：{result['h']}"
        )

        print(
            f"MCD support_fraction："
            f"{result['support_fraction']:.6f}"
        )

        print(
            f"原始MCD支持样本数："
            f"{int(np.sum(result['raw_support']))}"
        )

        print(
            f"最终MCD支持样本数："
            f"{int(np.sum(result['final_support']))}"
        )

        print(
            f"候选载荷阈值：{threshold:.6f}"
        )

        print(
            "\nPC解释方差："
        )

        for k in range(q):
            print(
                f"  PC{k+1:02d}："
                f"{result['explained_ratio'][k]*100:.4f}%"
                f"；累计="
                f"{result['cumulative_ratio'][k]*100:.4f}%"
            )

        print(
            "\nCLR全元素回映射检查："
        )

        print(
            f"  输入元素数：{D}"
        )

        print(
            f"  回映射元素数："
            f"{result['clr_loadings_all'].shape[0]}"
        )

        print(
            "  全部元素是否恢复：是"
        )

        print(
            "  CLR→ILR→CLR最大重构误差："
            f"{result['clr_reconstruction_error']:.3e}"
        )

        for k in range(q):
            print(
                f"  PC{k+1:02d} CLR载荷元素和："
                f"{np.sum(result['clr_loadings_all'][:, k]):.3e}"
            )

        if result[
            "warning_messages"
        ]:
            print(
                "\n运行警告："
            )

            for msg in result[
                "warning_messages"
            ]:
                print(
                    " - "
                    + msg
                )

        print(
            f"\n详细Excel结果：\n{excel_path}"
        )

        print(
            f"\n核心结果Excel：\n{核心excel_path}"
        )


        root = tk.Tk()
        root.withdraw()

        messagebox.showinfo(
            "计算完成",
            (
                "ILR + MCD-based Robust PCA 已完成。\n\n"
                f"样品数 n = {n}\n"
                f"输入元素数 D = {D}\n"
                f"ILR维数 p = {p}\n"
                f"输出 PC1-PC{q}\n"
                f"MCD alpha = {params['alpha']}\n"
                f"MCD本次实际随机种子 = {params['seed']}\n"
                f"自动计算 h = {result['h']}\n\n"
                f"CLR回映射已完整恢复全部 {D} 个原始元素。\n\n"
                f"详细Excel结果：\n{excel_path}\n\n"
                f"核心结果Excel：\n{核心excel_path}"
            ),
            parent=root,
        )

        root.destroy()


        try:
            if sys.platform.startswith(
                "win"
            ):
                os.startfile(
                    output_dir
                )
        except Exception:
            pass

    except SystemExit as exc:
        print(
            str(
                exc
            )
        )

    except Exception as exc:
        print(
            "\n"
            + "=" * 80
        )

        print(
            "程序运行失败"
        )

        print(
            "=" * 80
        )

        print(
            str(
                exc
            )
        )

        try:
            root = tk.Tk()
            root.withdraw()

            messagebox.showerror(
                "程序运行失败",
                str(
                    exc
                ),
                parent=root,
            )

            root.destroy()

        except Exception:
            pass


if __name__ == "__main__":
    main()
