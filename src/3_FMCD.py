# coding=utf-8

"""Compute robust Mahalanobis distances with FAST-MCD.

Input: Excel or CSV; two coordinate columns and preprocessed variables.
Options: MCD support, random starts, C-steps, cutoffs, and seed.
Output: Excel workbook with distances, anomaly classes, and MCD estimates.
Run: Start this script and follow the file and parameter dialogs."""

from __future__ import annotations

import os
import re
import math
import time
import traceback
from pathlib import Path
from datetime import datetime
from typing import List, Tuple, Optional

import numpy as np
import pandas as pd

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def _hidden_root():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def info(title: str, msg: str):
    root = _hidden_root()
    messagebox.showinfo(
        title,
        msg,
        parent=root,
    )
    root.destroy()


def warn(title: str, msg: str):
    root = _hidden_root()
    messagebox.showwarning(
        title,
        msg,
        parent=root,
    )
    root.destroy()


def ask_file_path(title: str) -> Path:
    root = _hidden_root()

    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel文件", "*.xlsx *.xls"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*"),
        ],
        parent=root,
    )

    root.destroy()

    if not file_path:
        raise SystemExit("未选择输入文件，程序结束。")

    return Path(file_path)


def ask_output_dir(default_name: str = "FMCD论文留档结果") -> Path:
    root = _hidden_root()

    parent = filedialog.askdirectory(
        title="请选择结果保存位置",
        parent=root,
    )

    root.destroy()

    if not parent:
        raise SystemExit("未选择输出位置，程序结束。")

    out_dir = (
        Path(parent)
        /
        f"{safe_name(default_name)}_{now_tag()}"
    )

    out_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    return out_dir


def ask_text(
    title: str,
    prompt: str,
    default: str = "",
) -> str:
    root = _hidden_root()

    value = simpledialog.askstring(
        title,
        prompt,
        initialvalue=default,
        parent=root,
    )

    root.destroy()

    if value is None:
        raise SystemExit("用户取消输入，程序结束。")

    return str(value).strip()


def ask_int(
    title: str,
    prompt: str,
    default: int,
    min_value: Optional[int] = None,
    max_value: Optional[int] = None,
) -> int:

    while True:
        text = ask_text(
            title,
            prompt,
            str(default),
        )

        try:
            value = int(text)

            if (
                min_value is not None
                and value < min_value
            ):
                raise ValueError(
                    f"必须 >= {min_value}"
                )

            if (
                max_value is not None
                and value > max_value
            ):
                raise ValueError(
                    f"必须 <= {max_value}"
                )

            return value

        except Exception as exc:
            warn(
                "输入错误",
                f"请输入整数。\n{exc}",
            )


def ask_float(
    title: str,
    prompt: str,
    default: float,
    min_value: Optional[float] = None,
    max_value: Optional[float] = None,
) -> float:

    while True:
        text = ask_text(
            title,
            prompt,
            str(default),
        )

        try:
            value = float(text)

            if (
                min_value is not None
                and value < min_value
            ):
                raise ValueError(
                    f"必须 >= {min_value}"
                )

            if (
                max_value is not None
                and value > max_value
            ):
                raise ValueError(
                    f"必须 <= {max_value}"
                )

            return value

        except Exception as exc:
            warn(
                "输入错误",
                f"请输入数字。\n{exc}",
            )


def ask_random_seed(
    default: int = 42,
) -> Optional[int]:

    text = ask_text(
        "随机种子",
        (
            "请输入随机种子。\n\n"
            "优选默认：42。\n"
            "固定随机种子可保证不同实验之间可复现。\n"
            "输入 None 表示每次随机。"
        ),
        str(default),
    )

    if text.lower() in {
        "none",
        "nan",
        "null",
        "",
    }:
        return None

    return int(text)


def now_tag() -> str:
    return datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )


def safe_name(name: str) -> str:
    name = str(name)

    name = re.sub(
        r'[\\/:*?"<>|\s]+',
        "_",
        name,
    )

    return (
        name.strip("_")
        or
        "result"
    )


def open_output_folder(path: Path):
    try:
        if os.name == "nt":
            os.startfile(
                str(path)
            )  # type: ignore[attr-defined]

        else:
            import subprocess

            subprocess.Popen(
                [
                    "xdg-open",
                    str(path),
                ]
            )

    except Exception:
        pass


def read_table(
    path: Path,
) -> Tuple[pd.DataFrame, str]:

    suffix = path.suffix.lower()

    if suffix == ".csv":

        errors = []

        for enc in [
            "utf-8-sig",
            "gb18030",
            "gbk",
            "utf-8",
        ]:
            try:
                df = pd.read_csv(
                    path,
                    encoding=enc,
                )

                return (
                    df,
                    "CSV",
                )

            except Exception as exc:
                errors.append(
                    f"{enc}: {exc}"
                )

        raise ValueError(
            "CSV文件编码无法识别。\n"
            "请优先另存为Excel .xlsx 后再运行。\n\n"
            +
            "\n".join(errors)
        )

    xls = pd.ExcelFile(path)

    if len(xls.sheet_names) == 1:
        sheet = xls.sheet_names[0]

    else:
        text = "\n".join(
            [
                f"{i}: {s}"
                for i, s
                in enumerate(
                    xls.sheet_names
                )
            ]
        )

        ans = ask_text(
            "选择Sheet",
            (
                "当前Excel包含多个Sheet：\n"
                f"{text}\n\n"
                "请输入Sheet序号或名称："
            ),
            "0",
        )

        if ans.isdigit():
            idx = int(ans)

            if (
                idx < 0
                or
                idx >= len(
                    xls.sheet_names
                )
            ):
                raise ValueError(
                    "Sheet序号超出范围。"
                )

            sheet = xls.sheet_names[idx]

        else:
            if ans not in xls.sheet_names:
                raise ValueError(
                    f"Sheet不存在：{ans}"
                )

            sheet = ans

    return (
        pd.read_excel(
            path,
            sheet_name=sheet,
        ),
        str(sheet),
    )


def split_geochem_table(df: pd.DataFrame):
    if df.shape[1] < 4:
        raise ValueError(
            "数据至少应包含："
            "2列坐标 + 至少2个FMCD变量。"
        )

    coord_cols = list(
        df.columns[:2]
    )

    element_cols = list(
        df.columns[2:]
    )

    coords = df.iloc[
        :,
        :2,
    ].copy()

    elements = df.iloc[
        :,
        2:,
    ].copy()

    return (
        coords,
        elements,
        coord_cols,
        element_cols,
    )


def parse_columns(
    text: str,
    available: List[str],
    default_all: bool = True,
) -> List[str]:

    if not text.strip():
        return (
            list(available)
            if default_all
            else []
        )

    parts = [
        p.strip()
        for p in re.split(
            r"[,，;；\s]+",
            text,
        )
        if p.strip()
    ]

    lower_map = {
        str(c).lower(): c
        for c in available
    }

    selected = []
    missing = []

    for p in parts:
        if p in available:
            selected.append(p)

        elif p.lower() in lower_map:
            selected.append(
                lower_map[
                    p.lower()
                ]
            )

        else:
            missing.append(p)

    if missing:
        raise ValueError(
            f"以下变量列不存在：{missing}\n"
            f"可用列：{available}"
        )


    return list(
        dict.fromkeys(
            selected
        )
    )


def validate_preprocessed_data(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """Check numeric inputs without applying any further transformation."""
    numeric = pd.DataFrame(
        index=df.index
    )

    problems = []

    for col in df.columns:
        original = df[col]

        converted = pd.to_numeric(
            original,
            errors="coerce",
        )

        non_numeric = (
            converted.isna()
            &
            ~original.isna()
        )

        if non_numeric.any():
            problems.append(
                f"{col}：存在{int(non_numeric.sum())}个非数值"
            )

        if converted.isna().any():
            problems.append(
                f"{col}：存在{int(converted.isna().sum())}个NaN/缺失值"
            )

        arr = converted.to_numpy(
            dtype=float
        )

        inf_count = int(
            np.isinf(arr).sum()
        )

        if inf_count > 0:
            problems.append(
                f"{col}：存在{inf_count}个Inf/-Inf"
            )

        numeric[col] = converted

    if problems:
        raise ValueError(
            "FMCD输入数据未通过校验。\n\n"
            "本程序默认数据已经完成预处理，"
            "不会自动替代、删除或变换任何值。\n\n"
            "请先修正：\n"
            +
            "\n".join(
                f"- {x}"
                for x in problems
            )
        )

    X = numeric.to_numpy(
        dtype=float
    )

    if not np.all(
        np.isfinite(X)
    ):
        raise ValueError(
            "参与FMCD的变量仍存在非有限数值。"
        )

    zero_variance = []

    for col in numeric.columns:
        arr = numeric[
            col
        ].to_numpy(
            dtype=float
        )

        if (
            len(arr) < 2
            or
            float(
                np.std(
                    arr,
                    ddof=1,
                )
            )
            <= 1e-15
        ):
            zero_variance.append(
                str(col)
            )

    if zero_variance:
        raise ValueError(
            "以下变量方差为0或接近0，不能进入FMCD：\n"
            +
            "、".join(
                zero_variance
            )
        )

    return numeric.astype(float)


def norm_ppf_acklam(
    p: float,
) -> float:
    """Approximate the inverse standard normal CDF."""
    if (
        p <= 0
        or
        p >= 1
    ):
        raise ValueError(
            "p must be in (0,1)"
        )

    a = [
        -3.969683028665376e+01,
        2.209460984245205e+02,
        -2.759285104469687e+02,
        1.383577518672690e+02,
        -3.066479806614716e+01,
        2.506628277459239e+00,
    ]

    b = [
        -5.447609879822406e+01,
        1.615858368580409e+02,
        -1.556989798598866e+02,
        6.680131188771972e+01,
        -1.328068155288572e+01,
    ]

    c = [
        -7.784894002430293e-03,
        -3.223964580411365e-01,
        -2.400758277161838e+00,
        -2.549732539343734e+00,
        4.374664141464968e+00,
        2.938163982698783e+00,
    ]

    d = [
        7.784695709041462e-03,
        3.224671290700398e-01,
        2.445134137142996e+00,
        3.754408661907416e+00,
    ]

    plow = 0.02425
    phigh = 1 - plow

    if p < plow:
        q = math.sqrt(
            -2
            *
            math.log(p)
        )

        return (
            (
                (
                    (
                        (
                            c[0] * q
                            +
                            c[1]
                        )
                        * q
                        +
                        c[2]
                    )
                    * q
                    +
                    c[3]
                )
                * q
                +
                c[4]
            )
            * q
            +
            c[5]
        ) / (
            (
                (
                    (
                        d[0] * q
                        +
                        d[1]
                    )
                    * q
                    +
                    d[2]
                )
                * q
                +
                d[3]
            )
            * q
            +
            1
        )

    if p > phigh:
        q = math.sqrt(
            -2
            *
            math.log(
                1 - p
            )
        )

        return -(
            (
                (
                    (
                        (
                            c[0] * q
                            +
                            c[1]
                        )
                        * q
                        +
                        c[2]
                    )
                    * q
                    +
                    c[3]
                )
                * q
                +
                c[4]
            )
            * q
            +
            c[5]
        ) / (
            (
                (
                    (
                        d[0] * q
                        +
                        d[1]
                    )
                    * q
                    +
                    d[2]
                )
                * q
                +
                d[3]
            )
            * q
            +
            1
        )

    q = p - 0.5
    r = q * q

    return (
        (
            (
                (
                    (
                        (
                            a[0] * r
                            +
                            a[1]
                        )
                        * r
                        +
                        a[2]
                    )
                    * r
                    +
                    a[3]
                )
                * r
                +
                a[4]
            )
            * r
            +
            a[5]
        )
        *
        q
    ) / (
        (
            (
                (
                    (
                        b[0] * r
                        +
                        b[1]
                    )
                    * r
                    +
                    b[2]
                )
                * r
                +
                b[3]
            )
            * r
            +
            b[4]
        )
        * r
        +
        1
    )


def chi2_ppf(
    q: float,
    df: int,
) -> float:

    try:
        from scipy.stats import chi2  # type: ignore

        return float(
            chi2.ppf(
                q,
                df=df,
            )
        )

    except Exception:

        z = norm_ppf_acklam(q)

        return float(
            df
            *
            (
                1
                -
                2
                /
                (
                    9
                    *
                    df
                )
                +
                z
                *
                math.sqrt(
                    2
                    /
                    (
                        9
                        *
                        df
                    )
                )
            )
            ** 3
        )


def _center_cov_mle(
    X: np.ndarray,
) -> Tuple[
    np.ndarray,
    np.ndarray,
]:

    center = X.mean(
        axis=0
    )

    diff = (
        X
        -
        center
    )

    cov = (
        diff.T
        @
        diff
    ) / max(
        len(X),
        1,
    )

    return (
        center,
        cov,
    )


def _center_cov_unbiased(
    X: np.ndarray,
) -> Tuple[
    np.ndarray,
    np.ndarray,
]:

    center = X.mean(
        axis=0
    )

    if len(X) <= 1:
        cov = np.zeros(
            (
                X.shape[1],
                X.shape[1],
            )
        )

    else:
        cov = np.cov(
            X,
            rowvar=False,
            ddof=1,
        )

    return (
        center,
        np.atleast_2d(cov),
    )


def _safe_logdet(
    cov: np.ndarray,
    ridge: float = 0.0,
) -> Tuple[
    float,
    int,
    bool,
]:

    C = np.array(
        cov,
        dtype=float,
        copy=True,
    )

    if ridge > 0:
        C = (
            C
            +
            np.eye(
                C.shape[0]
            )
            *
            ridge
        )

    rank = int(
        np.linalg.matrix_rank(
            C,
            tol=1e-10,
        )
    )

    sign, ld = np.linalg.slogdet(C)

    if (
        sign <= 0
        or
        not np.isfinite(ld)
    ):
        if rank < C.shape[0]:
            return (
                float("-inf"),
                rank,
                True,
            )

        return (
            float("inf"),
            rank,
            False,
        )

    return (
        float(ld),
        rank,
        False,
    )


def mahalanobis2(
    X: np.ndarray,
    center: np.ndarray,
    cov: np.ndarray,
    ridge: float = 0.0,
) -> np.ndarray:

    C = np.asarray(
        cov,
        dtype=float,
    )

    if ridge > 0:
        C = (
            C
            +
            np.eye(
                C.shape[0]
            )
            *
            ridge
        )

    inv = np.linalg.pinv(C)

    diff = (
        X
        -
        center
    )

    d2 = np.einsum(
        "ij,jk,ik->i",
        diff,
        inv,
        diff,
    )

    return np.maximum(
        d2,
        0.0,
    )


def _c_step(
    X: np.ndarray,
    center: np.ndarray,
    cov: np.ndarray,
    h: int,
    ridge: float = 1e-12,
):
    d2 = mahalanobis2(
        X,
        center,
        cov,
        ridge=ridge,
    )

    idx = np.argpartition(
        d2,
        h - 1,
    )[:h]

    idx = idx[
        np.argsort(
            d2[idx]
        )
    ]

    (
        new_center,
        new_cov,
    ) = _center_cov_mle(
        X[idx]
    )

    (
        logdet,
        rank,
        singular,
    ) = _safe_logdet(
        new_cov
    )

    return (
        idx,
        new_center,
        new_cov,
        logdet,
        rank,
        singular,
        d2,
    )


def _initial_h_from_p_plus_1(
    X: np.ndarray,
    h: int,
    rng: np.random.Generator,
    ridge: float = 1e-12,
) -> np.ndarray:

    n, p = X.shape

    if n < p + 1:
        raise ValueError(
            "样品数 n 必须至少大于 p+1。"
        )

    available = np.arange(n)

    J = rng.choice(
        available,
        size=p + 1,
        replace=False,
    ).tolist()

    remaining = [
        i
        for i in available.tolist()
        if i not in J
    ]

    while True:
        (
            center0,
            cov0,
        ) = _center_cov_mle(
            X[J]
        )

        (
            _,
            rank,
            _,
        ) = _safe_logdet(
            cov0
        )

        if (
            rank >= p
            or
            len(J) >= n
        ):
            break

        if not remaining:
            break

        add_idx = rng.choice(
            np.array(remaining),
            size=1,
            replace=False,
        )[0]

        J.append(
            int(add_idx)
        )

        remaining.remove(
            int(add_idx)
        )

    d2 = mahalanobis2(
        X,
        center0,
        cov0,
        ridge=ridge,
    )

    H = np.argpartition(
        d2,
        h - 1,
    )[:h]

    H = H[
        np.argsort(
            d2[H]
        )
    ]

    return H


def _iterate_c_steps(
    X: np.ndarray,
    H: np.ndarray,
    h: int,
    max_steps: int = 100,
    tol: float = 1e-12,
    ridge: float = 1e-12,
):
    """Iterate C-steps until convergence or the iteration limit."""
    H = np.array(
        H,
        dtype=int,
    )

    (
        center,
        cov,
    ) = _center_cov_mle(
        X[H]
    )

    (
        old_logdet,
        _,
        _,
    ) = _safe_logdet(
        cov
    )

    n_steps = 0

    for step in range(
        1,
        max_steps + 1,
    ):
        n_steps = step

        (
            H_new,
            center_new,
            cov_new,
            logdet_new,
            _,
            _,
            _,
        ) = _c_step(
            X,
            center,
            cov,
            h,
            ridge=ridge,
        )

        same_subset = np.array_equal(
            np.sort(H),
            np.sort(H_new),
        )

        no_decrease = (
            (
                logdet_new
                >=
                old_logdet
                -
                tol
            )
            if (
                np.isfinite(old_logdet)
                and
                np.isfinite(logdet_new)
            )
            else False
        )

        H = H_new
        center = center_new
        cov = cov_new

        if logdet_new == float("-inf"):
            break

        if (
            same_subset
            or
            no_decrease
        ):
            break

        old_logdet = logdet_new

    (
        final_logdet,
        final_rank,
        final_singular,
    ) = _safe_logdet(
        cov
    )

    return {
        "H": H,
        "center": center,
        "cov": cov,
        "logdet": final_logdet,
        "rank": final_rank,
        "singular": final_singular,
        "n_steps": n_steps,
    }


def _deduplicate_candidates(
    candidates: List[dict],
) -> List[dict]:

    seen = set()
    unique = []

    for c in sorted(
        candidates,
        key=lambda x: x["logdet"],
    ):
        key = tuple(
            np.sort(
                c["H"]
            ).tolist()
        )

        if key not in seen:
            unique.append(c)
            seen.add(key)

    return unique


def fast_mcd_small_sample(
    X: np.ndarray,
    h: int,
    starts: int = 1000,
    initial_c_steps: int = 2,
    keep_best: int = 20,
    max_c_steps: int = 100,
    seed: Optional[int] = 42,
    ridge: float = 1e-12,
):
    """Find an MCD subset using random starts and C-steps."""
    rng = np.random.default_rng(seed)

    n, p = X.shape

    candidates = []

    t0 = time.time()

    for _ in range(starts):
        H0 = _initial_h_from_p_plus_1(
            X,
            h,
            rng,
            ridge=ridge,
        )

        res0 = _iterate_c_steps(
            X,
            H0,
            h,
            max_steps=initial_c_steps,
            ridge=ridge,
        )

        candidates.append(res0)

    candidates = _deduplicate_candidates(
        candidates
    )

    if not candidates:
        raise RuntimeError(
            "未得到有效FAST-MCD初始候选解。"
        )

    best_initial = sorted(
        candidates,
        key=lambda x: x["logdet"],
    )[:keep_best]

    final_candidates = []

    for cand in best_initial:
        res = _iterate_c_steps(
            X,
            cand["H"],
            h,
            max_steps=max_c_steps,
            ridge=ridge,
        )

        final_candidates.append(res)

    final_candidates = _deduplicate_candidates(
        final_candidates
    )

    if not final_candidates:
        raise RuntimeError(
            "未得到有效FAST-MCD最终候选解。"
        )

    best = sorted(
        final_candidates,
        key=lambda x: x["logdet"],
    )[0]

    info_dict = {
        "algorithm": (
            "FAST-MCD core: random(p+1) starts + "
            "selective C-step + convergence"
        ),
        "n": int(n),
        "p": int(p),
        "h": int(h),
        "starts": int(starts),
        "initial_c_steps": int(initial_c_steps),
        "keep_best": int(keep_best),
        "max_c_steps": int(max_c_steps),
        "raw_logdet": float(
            best["logdet"]
        ),
        "raw_rank": int(
            best["rank"]
        ),
        "raw_singular": bool(
            best["singular"]
        ),
        "raw_support_size": int(
            len(best["H"])
        ),
        "final_c_steps": int(
            best["n_steps"]
        ),
        "runtime_seconds": float(
            time.time()
            -
            t0
        ),
        "seed": seed,
    }

    return (
        best,
        info_dict,
    )


def fast_mcd(
    X: np.ndarray,
    support_fraction: float = 0.75,
    starts: int = 1000,
    initial_c_steps: int = 2,
    keep_best: int = 20,
    max_c_steps: int = 100,
    seed: Optional[int] = 42,
    ridge: float = 1e-12,
    reweight_quantile: float = 0.975,
):
    X = np.asarray(
        X,
        dtype=float,
    )

    if X.ndim != 2:
        raise ValueError(
            "X必须为二维矩阵。"
        )

    n, p = X.shape

    if p < 2:
        raise ValueError(
            "FMCD至少需要2个变量。"
        )

    if n <= p + 1:
        raise ValueError(
            "样品数必须大于 p+1。"
        )

    h = int(
        math.floor(
            n
            *
            support_fraction
        )
    )

    h_min = int(
        math.ceil(
            (
                n
                +
                p
                +
                1
            )
            /
            2.0
        )
    )

    if h < h_min:
        raise ValueError(
            f"support_fraction过低：h={h}，"
            f"理论下限h_min={h_min}。\n"
            f"support_fraction至少约为"
            f"{h_min / n:.4f}。"
        )

    h = min(
        h,
        n,
    )

    if h == n:
        (
            raw_center,
            raw_cov,
        ) = _center_cov_mle(
            X
        )

        (
            raw_logdet,
            raw_rank,
            raw_singular,
        ) = _safe_logdet(
            raw_cov
        )

        best = {
            "H": np.arange(n),
            "center": raw_center,
            "cov": raw_cov,
            "logdet": raw_logdet,
            "rank": raw_rank,
            "singular": raw_singular,
            "n_steps": 0,
        }

        info_dict = {
            "algorithm": (
                "h=n; MCD退化为经典均值/协方差"
            ),
            "n": int(n),
            "p": int(p),
            "h": int(h),
            "starts": int(starts),
            "initial_c_steps": int(
                initial_c_steps
            ),
            "keep_best": int(
                keep_best
            ),
            "max_c_steps": int(
                max_c_steps
            ),
            "raw_logdet": float(
                raw_logdet
            ),
            "raw_rank": int(
                raw_rank
            ),
            "raw_singular": bool(
                raw_singular
            ),
            "raw_support_size": int(n),
            "final_c_steps": 0,
            "runtime_seconds": 0.0,
            "seed": seed,
        }

    else:
        if n > 600:
            warn(
                "样品数较大",
                (
                    "当前 n > 600。\n"
                    "程序继续运行 random(p+1) + "
                    "selective C-step 核心流程，"
                    "运行时间可能增加。"
                ),
            )

        (
            best,
            info_dict,
        ) = fast_mcd_small_sample(
            X,
            h=h,
            starts=starts,
            initial_c_steps=initial_c_steps,
            keep_best=keep_best,
            max_c_steps=max_c_steps,
            seed=seed,
            ridge=ridge,
        )

    raw_center = best[
        "center"
    ]

    raw_cov = best[
        "cov"
    ]


    raw_md2 = mahalanobis2(
        X,
        raw_center,
        raw_cov,
        ridge=ridge,
    )

    chi2_median = chi2_ppf(
        0.5,
        p,
    )

    raw_md2_median = float(
        np.median(
            raw_md2
        )
    )

    consistency_factor = (
        raw_md2_median
        /
        chi2_median
        if (
            chi2_median > 0
            and
            np.isfinite(
                raw_md2_median
            )
        )
        else 1.0
    )

    if (
        not np.isfinite(
            consistency_factor
        )
        or
        consistency_factor <= 0
    ):
        consistency_factor = 1.0

    # Apply median-based covariance scaling before reweighting.
    consistent_cov = (
        raw_cov
        *
        consistency_factor
    )

    consistent_md2 = mahalanobis2(
        X,
        raw_center,
        consistent_cov,
        ridge=ridge,
    )


    reweight_cutoff = chi2_ppf(
        reweight_quantile,
        p,
    )

    weights = (
        consistent_md2
        <=
        reweight_cutoff
    ).astype(int)

    if weights.sum() <= p + 1:
        reweighted_center = (
            raw_center.copy()
        )

        reweighted_cov = (
            consistent_cov.copy()
        )

        reweight_note = (
            "再加权支持样品数过少，"
            "使用consistency-corrected raw MCD。"
        )

    else:
        Xw = X[
            weights == 1
        ]

        (
            reweighted_center,
            reweighted_cov,
        ) = _center_cov_unbiased(
            Xw
        )

        reweight_note = (
            "完成one-step reweighted MCD。"
        )

    # Evaluate distances using the final reweighted estimate.
    final_md2 = mahalanobis2(
        X,
        reweighted_center,
        reweighted_cov,
        ridge=ridge,
    )

    final_cov_rank = int(
        np.linalg.matrix_rank(
            reweighted_cov,
            tol=1e-10,
        )
    )

    try:
        final_cov_condition = float(
            np.linalg.cond(
                reweighted_cov
            )
        )
    except Exception:
        final_cov_condition = float(
            "nan"
        )

    info_dict.update({
        "support_fraction": float(
            support_fraction
        ),
        "h": int(h),
        "h_min": int(
            h_min
        ),
        "breakdown_value_approx": float(
            (
                n
                -
                h
                +
                1
            )
            /
            n
        ),
        "ridge": float(
            ridge
        ),
        "chi2_median": float(
            chi2_median
        ),
        "raw_MD2_median": float(
            raw_md2_median
        ),
        "consistency_factor": float(
            consistency_factor
        ),
        "reweight_quantile": float(
            reweight_quantile
        ),
        "reweight_cutoff_MD2": float(
            reweight_cutoff
        ),
        "reweighted_support_size": int(
            weights.sum()
        ),
        "reweighted_support_ratio": float(
            weights.sum()
            /
            n
        ),
        "reweight_note": (
            reweight_note
        ),
        "final_cov_rank": int(
            final_cov_rank
        ),
        "final_cov_condition_number": float(
            final_cov_condition
        ),
    })

    return {
        "final_md2": final_md2,
        "reweighted_center": reweighted_center,
        "reweighted_cov": reweighted_cov,
        "reweighted_support": weights,
        "info": info_dict,
    }


def write_result_excel(
    path: Path,
    sheets: dict[str, pd.DataFrame],
):
    """Write MCD results to a formatted Excel workbook."""
    with pd.ExcelWriter(
        path,
        engine="openpyxl",
    ) as writer:

        for sheet_name, df in sheets.items():
            df.to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False,
            )


    from openpyxl import load_workbook
    from openpyxl.styles import (
        Alignment,
        Font,
        PatternFill,
        Border,
        Side,
    )
    from openpyxl.utils import (
        get_column_letter,
    )

    wb = load_workbook(
        path
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        name="Microsoft YaHei",
        color="FFFFFF",
        bold=True,
        size=10,
    )

    body_font = Font(
        name="Microsoft YaHei",
        size=10,
    )

    thin_gray = Side(
        style="thin",
        color="D9E1F2",
    )

    border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray,
    )

    level_fill = {
        "背景": "FFFFFF",
        "弱异常": "FFF2CC",
        "中异常": "FCE4D6",
        "强异常": "F4CCCC",
    }

    for ws in wb.worksheets:
        ws.freeze_panes = (
            "A2"
        )

        if ws.max_row >= 1:
            ws.auto_filter.ref = (
                ws.dimensions
            )


        for cell in ws[1]:
            cell.fill = (
                header_fill
            )

            cell.font = (
                header_font
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            cell.border = (
                border
            )


        for row in ws.iter_rows(
            min_row=2
        ):
            for cell in row:
                cell.font = (
                    body_font
                )

                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=True,
                )

                cell.border = (
                    border
                )


        level_col = None

        for cell in ws[1]:
            if cell.value == "FMCD_异常等级":
                level_col = (
                    cell.column
                )
                break

        if level_col is not None:
            for r in range(
                2,
                ws.max_row + 1,
            ):
                level = ws.cell(
                    row=r,
                    column=level_col,
                ).value

                if level in level_fill:
                    ws.cell(
                        row=r,
                        column=level_col,
                    ).fill = PatternFill(
                        fill_type="solid",
                        fgColor=(
                            level_fill[
                                level
                            ]
                        ),
                    )


        for col_idx in range(
            1,
            ws.max_column + 1,
        ):
            max_len = 10

            for row_idx in range(
                1,
                min(
                    ws.max_row,
                    300,
                )
                +
                1,
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
                                str(value)
                            )
                            +
                            2,
                            48,
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
        path
    )


def main():
    info(
        "FMCD / FAST-MCD论文留档版",
        (
            "本版本默认输入数据已经完成全部必要预处理。\n\n"
            "不会再做log10、Z-score、缺失值替代、0值替代、"
            "ILR/CLR或异常值删除。\n\n"
            "最终只保留论文和复现实验真正需要的FMCD结果与关键参数。"
        ),
    )


    in_path = ask_file_path(
        "请选择已经完成预处理的FMCD输入数据表"
    )

    (
        df,
        sheet_name,
    ) = read_table(
        in_path
    )

    (
        coords_raw,
        elements_raw,
        coord_cols,
        element_cols,
    ) = split_geochem_table(
        df
    )


    elem_text = ask_text(
        "选择参与FMCD的变量",
        (
            "请输入参与FMCD的变量名，用逗号分隔。\n\n"
            "例如：ILR1,ILR2,ILR3,ILR4\n"
            "或输入其他已经完成预处理的变量。\n\n"
            "直接回车默认使用第三列以后全部变量。\n\n"
            "本程序不会再次进行任何预处理。"
        ),
        "",
    )

    selected = parse_columns(
        elem_text,
        element_cols,
        default_all=True,
    )

    if len(selected) < 2:
        raise ValueError(
            "至少选择2个FMCD变量。"
        )


    M_df = validate_preprocessed_data(
        elements_raw[
            selected
        ].copy()
    )

    X = M_df.to_numpy(
        dtype=float
    )

    n, p = X.shape

    if n <= p + 1:
        raise ValueError(
            f"样品数 n={n} 必须大于 p+1={p+1}。"
        )


    out_dir = ask_output_dir(
        "FMCD论文留档结果"
    )


    support_fraction = ask_float(
        "support_fraction / h比例",
        (
            "请输入 support_fraction = h/n。\n\n"
            "优选默认：0.75\n"
            "兼顾稳健性与统计效率。"
        ),
        0.75,
        min_value=0.5,
        max_value=1.0,
    )

    starts = ask_int(
        "随机(p+1)初始子集次数",
        (
            "请输入随机初始子集次数 starts。\n\n"
            "优选默认：1000\n"
        "用于提高随机初始子集搜索稳定性。"
        ),
        1000,
        min_value=100,
    )

    initial_c_steps = ask_int(
        "初筛C-step次数",
        (
            "请输入初筛阶段每个初始子集执行的C-step次数。\n\n"
            "优选默认：2\n"
            "符合selective C-step的常用设置。"
        ),
        2,
        min_value=1,
    )

    keep_best = ask_int(
        "保留最优候选数",
        (
            "请输入初筛后保留的最优候选数量。\n\n"
            "优选默认：20\n"
        "在初筛后保留更多优质候选，降低偶然初始解影响。"
        ),
        20,
        min_value=1,
    )

    max_c_steps = ask_int(
        "最终C-step最大次数",
        (
            "请输入最终阶段C-step最大次数。\n"
            "收敛后会提前停止。\n\n"
            "优选默认：100\n"
            "通常足以收敛，程序会提前停止。"
        ),
        100,
        min_value=5,
    )

    reweight_q = ask_float(
        "再加权阈值分位数",
        (
            "请输入one-step reweighting使用的卡方分位数。\n\n"
            "优选默认：0.975\n"
            "兼顾稳健背景保留和异常排除。"
        ),
        0.975,
        min_value=0.5,
        max_value=0.9999,
    )

    anomaly_q = ask_float(
        "异常判别阈值分位数",
        (
            "请输入最终FMCD异常判别使用的卡方分位数。\n\n"
            "优选默认：0.975\n"
            "适合弱异常识别；若需更严格判别可改为0.99。"
        ),
        0.975,
        min_value=0.5,
        max_value=0.9999,
    )

    seed = ask_random_seed(
        42
    )

    ridge = 1e-12


    total_t0 = time.time()

    result_mcd = fast_mcd(
        X,
        support_fraction=(
            support_fraction
        ),
        starts=starts,
        initial_c_steps=(
            initial_c_steps
        ),
        keep_best=keep_best,
        max_c_steps=(
            max_c_steps
        ),
        seed=seed,
        ridge=ridge,
        reweight_quantile=(
            reweight_q
        ),
    )

    total_runtime = float(
        time.time()
        -
        total_t0
    )

    final_md2 = np.asarray(
        result_mcd[
            "final_md2"
        ],
        dtype=float,
    )

    final_md = np.sqrt(
        np.maximum(
            final_md2,
            0.0,
        )
    )


    cutoff_md2 = chi2_ppf(
        anomaly_q,
        p,
    )

    cutoff_md = math.sqrt(
        cutoff_md2
    )

    anomaly_index = (
        final_md2
        /
        cutoff_md2
    )

    anomaly_flag = np.where(
        final_md2
        >
        cutoff_md2,
        "异常",
        "背景",
    )

    anomaly_level = pd.cut(
        anomaly_index,
        bins=[
            -np.inf,
            1.0,
            1.5,
            2.0,
            np.inf,
        ],
        labels=[
            "背景",
            "弱异常",
            "中异常",
            "强异常",
        ],
        include_lowest=True,
        right=True,
    ).astype(str)


    anomaly_numeric_map = {
        "背景": 0,
        "弱异常": 1,
        "中异常": 2,
        "强异常": 3,
    }

    anomaly_numeric = (
        pd.Series(anomaly_level)
        .map(anomaly_numeric_map)
        .astype(int)
        .to_numpy()
    )


    result_df = pd.DataFrame({
        "原始行号": (
            np.arange(
                len(df)
            )
            +
            2
        ),
        str(
            coord_cols[
                0
            ]
        ): coords_raw.iloc[
            :,
            0
        ].to_numpy(),
        str(
            coord_cols[
                1
            ]
        ): coords_raw.iloc[
            :,
            1
        ].to_numpy(),
        "FMCD_MD2": final_md2,
        "FMCD_MD": final_md,
        "异常阈值_MD2": np.full(
            n,
            cutoff_md2,
            dtype=float,
        ),
        "FMCD_异常指数": (
            anomaly_index
        ),
        "FMCD_异常标记": (
            anomaly_flag
        ),
        "FMCD_异常等级": (
            anomaly_level
        ),
        "异常数字化": (
            anomaly_numeric
        ),
    })


    level_order = [
        "背景",
        "弱异常",
        "中异常",
        "强异常",
    ]

    counts_series = pd.Series(
        anomaly_level
    ).value_counts()

    anomaly_summary_rows = []

    for level in level_order:
        count = int(
            counts_series.get(
                level,
                0,
            )
        )

        anomaly_summary_rows.append({
            "异常等级": level,
            "样品数": count,
            "比例": (
                count
                /
                n
            ),
            "判定范围": {
                "背景": "FMCD_异常指数 ≤ 1.0",
                "弱异常": "1.0 < FMCD_异常指数 ≤ 1.5",
                "中异常": "1.5 < FMCD_异常指数 ≤ 2.0",
                "强异常": "FMCD_异常指数 > 2.0",
            }[
                level
            ],
        })

    anomaly_summary_df = pd.DataFrame(
        anomaly_summary_rows
    )

    abnormal_count = int(
        np.sum(
            final_md2
            >
            cutoff_md2
        )
    )

    abnormal_ratio = (
        abnormal_count
        /
        n
    )


    info_mcd = result_mcd[
        "info"
    ]

    param_rows = [

        (
            "数据",
            "输入文件",
            str(
                in_path
            ),
            "本次FMCD输入文件",
        ),
        (
            "数据",
            "Sheet",
            sheet_name,
            "本次读取的工作表",
        ),
        (
            "数据",
            "输入数据状态",
            "已完成预处理；FMCD程序未再次进行任何数据变换",
            "用于论文方法说明",
        ),
        (
            "数据",
            "坐标列",
            "、".join(
                map(
                    str,
                    coord_cols,
                )
            ),
            "结果表中保留的空间坐标",
        ),
        (
            "数据",
            "参与FMCD变量",
            "、".join(
                map(
                    str,
                    M_df.columns,
                )
            ),
            "用于计算稳健马氏距离的变量",
        ),
        (
            "数据",
            "样品数_n",
            n,
            "进入FMCD的样品数",
        ),
        (
            "数据",
            "变量数_p",
            p,
            "FMCD空间维数",
        ),
        (
            "数据",
            "卡方自由度_df",
            p,
            "异常阈值使用χ²(p)",
        ),


        (
            "FAST-MCD",
            "support_fraction",
            support_fraction,
            "h/n，稳健支持子集比例",
        ),
        (
            "FAST-MCD",
            "h",
            info_mcd[
                "h"
            ],
            "MCD稳健支持子集样品数",
        ),
        (
            "FAST-MCD",
            "h_min",
            info_mcd[
                "h_min"
            ],
            "ceil((n+p+1)/2)",
        ),
        (
            "FAST-MCD",
            "breakdown_value_approx",
            info_mcd[
                "breakdown_value_approx"
            ],
            "(n-h+1)/n",
        ),
        (
            "FAST-MCD",
            "starts",
            starts,
            "随机(p+1)初始子集数量",
        ),
        (
            "FAST-MCD",
            "initial_c_steps",
            initial_c_steps,
            "初筛阶段C-step次数",
        ),
        (
            "FAST-MCD",
            "keep_best",
            keep_best,
            "初筛后继续迭代的最优候选数量",
        ),
        (
            "FAST-MCD",
            "max_c_steps",
            max_c_steps,
            "最终阶段最大C-step次数",
        ),
        (
            "FAST-MCD",
            "final_c_steps",
            info_mcd.get(
                "final_c_steps"
            ),
            "最终最优解实际C-step次数",
        ),
        (
            "FAST-MCD",
            "ridge",
            ridge,
            "数值稳定用微小对角正则项",
        ),
        (
            "FAST-MCD",
            "random_seed",
            seed,
            "随机种子，固定后可复现",
        ),


        (
            "Consistency correction",
            "chi2_median",
            info_mcd[
                "chi2_median"
            ],
            "χ²(p)中位数",
        ),
        (
            "Consistency correction",
            "raw_MD2_median",
            info_mcd[
                "raw_MD2_median"
            ],
            "Raw MCD稳健距离平方中位数",
        ),
        (
            "Consistency correction",
            "consistency_factor",
            info_mcd[
                "consistency_factor"
            ],
            "Raw协方差一致性修正因子",
        ),


        (
            "One-step reweighting",
            "reweight_quantile",
            reweight_q,
            "再加权使用的χ²分位数",
        ),
        (
            "One-step reweighting",
            "reweight_cutoff_MD2",
            info_mcd[
                "reweight_cutoff_MD2"
            ],
            "再加权距离平方阈值",
        ),
        (
            "One-step reweighting",
            "reweighted_support_size",
            info_mcd[
                "reweighted_support_size"
            ],
            "最终再加权支持样品数",
        ),
        (
            "One-step reweighting",
            "reweighted_support_ratio",
            info_mcd[
                "reweighted_support_ratio"
            ],
            "最终再加权支持样品比例",
        ),
        (
            "One-step reweighting",
            "reweight_note",
            info_mcd[
                "reweight_note"
            ],
            "再加权执行状态",
        ),


        (
            "异常判别",
            "anomaly_quantile",
            anomaly_q,
            "最终异常判别χ²分位数",
        ),
        (
            "异常判别",
            "anomaly_cutoff_MD2",
            cutoff_md2,
            "最终FMCD_MD2异常阈值",
        ),
        (
            "异常判别",
            "anomaly_cutoff_MD",
            cutoff_md,
            "最终FMCD_MD异常阈值",
        ),
        (
            "异常判别",
            "异常样品数",
            abnormal_count,
            "FMCD_MD2 > anomaly_cutoff_MD2",
        ),
        (
            "异常判别",
            "异常样品比例",
            abnormal_ratio,
            "异常样品数 / n",
        ),
        (
            "异常判别",
            "背景样品数",
            int(
                counts_series.get(
                    "背景",
                    0,
                )
            ),
            "FMCD_异常指数≤1",
        ),
        (
            "异常判别",
            "弱异常样品数",
            int(
                counts_series.get(
                    "弱异常",
                    0,
                )
            ),
            "1<指数≤1.5",
        ),
        (
            "异常判别",
            "中异常样品数",
            int(
                counts_series.get(
                    "中异常",
                    0,
                )
            ),
            "1.5<指数≤2",
        ),
        (
            "异常判别",
            "强异常样品数",
            int(
                counts_series.get(
                    "强异常",
                    0,
                )
            ),
            "指数>2",
        ),


        (
            "MCD诊断",
            "raw_logdet",
            info_mcd.get(
                "raw_logdet"
            ),
            "Raw MCD协方差矩阵log(det)",
        ),
        (
            "MCD诊断",
            "raw_rank",
            info_mcd.get(
                "raw_rank"
            ),
            "Raw MCD协方差矩阵秩",
        ),
        (
            "MCD诊断",
            "raw_singular",
            info_mcd.get(
                "raw_singular"
            ),
            "Raw MCD协方差是否奇异",
        ),
        (
            "MCD诊断",
            "raw_support_size",
            info_mcd.get(
                "raw_support_size"
            ),
            "Raw MCD支持子集大小",
        ),
        (
            "MCD诊断",
            "final_cov_rank",
            info_mcd[
                "final_cov_rank"
            ],
            "最终稳健协方差矩阵秩",
        ),
        (
            "MCD诊断",
            "final_cov_condition_number",
            info_mcd[
                "final_cov_condition_number"
            ],
            "最终稳健协方差矩阵条件数",
        ),


        (
            "运行",
            "algorithm",
            info_mcd.get(
                "algorithm"
            ),
            "FAST-MCD核心流程说明",
        ),
        (
            "运行",
            "总运行时间_秒",
            total_runtime,
            "本次完整FMCD计算耗时",
        ),
        (
            "结果定义",
            "FMCD_MD2",
            "Reweighted FAST-MCD robust Mahalanobis distance squared",
            "论文主要FMCD距离指标",
        ),
        (
            "结果定义",
            "FMCD_异常指数",
            "FMCD_MD2 / anomaly_cutoff_MD2",
            "用于标准化不同模型/分组的异常强度",
        ),
        (
            "结果定义",
            "异常数字化",
            "背景=0；弱异常=1；中异常=2；强异常=3",
            "用于后续BPNN、统计分析、GIS制图或分类编码",
        ),
        (
            "参数说明",
            "默认参数策略",
            "面向分组后少变量FMCD弱异常识别的稳健优选默认值",
            "默认值兼顾稳健性、搜索稳定性、计算效率和实验复现性，并非数据无关的理论唯一最优",
        ),
    ]

    param_df = pd.DataFrame(
        param_rows,
        columns=[
            "参数类别",
            "关键参数",
            "取值",
            "说明",
        ],
    )


    robust_center_df = pd.DataFrame({
        "变量": M_df.columns,
        "最终稳健中心": result_mcd[
            "reweighted_center"
        ],
    })


    robust_cov_df = pd.DataFrame(
        result_mcd[
            "reweighted_cov"
        ],
        columns=M_df.columns,
    )

    robust_cov_df.insert(
        0,
        "变量",
        M_df.columns,
    )


    out_xlsx = (
        out_dir
        /
        (
            f"{safe_name(in_path.stem)}_"
            f"FMCD论文留档结果_"
            f"{now_tag()}.xlsx"
        )
    )

    write_result_excel(
        path=out_xlsx,
        sheets={
            "01_FMCD结果": result_df,
            "02_关键参数": param_df,
            "03_异常统计": anomaly_summary_df,
            "04_最终稳健中心": robust_center_df,
            "05_最终稳健协方差": robust_cov_df,
        },
    )


    print(
        "\n"
        +
        "=" * 80
    )

    print(
        "FMCD / FAST-MCD 计算完成"
    )

    print(
        "=" * 80
    )

    print(
        f"样品数 n：{n}"
    )

    print(
        f"变量数 p：{p}"
    )

    print(
        "参与变量："
        +
        "、".join(
            map(
                str,
                M_df.columns,
            )
        )
    )

    print(
        f"support_fraction：{support_fraction}"
    )

    print(
        f"h：{info_mcd['h']}"
    )

    print(
        f"breakdown value≈{info_mcd['breakdown_value_approx']:.4f}"
    )

    print(
        f"再加权支持样品数：{info_mcd['reweighted_support_size']}"
    )

    print(
        f"异常分位数：{anomaly_q}"
    )

    print(
        f"异常阈值 MD²：{cutoff_md2:.6g}"
    )

    print(
        f"异常样品数：{abnormal_count}"
    )

    print(
        f"异常比例：{abnormal_ratio:.4%}"
    )


    print(
        "异常数字化：背景=0，弱异常=1，中异常=2，强异常=3"
    )

    print(
        f"运行时间：{total_runtime:.3f} 秒"
    )

    print(
        f"\n结果文件：\n{out_xlsx}"
    )

    info(
        "完成",
        (
            "FMCD / FAST-MCD计算完成。\n\n"
            f"样品数：{n}\n"
            f"变量数：{p}\n"
            f"异常样品数：{abnormal_count}\n"
            f"异常比例：{abnormal_ratio:.2%}\n\n"
            "Excel包含：\n"
            "01_FMCD结果\n"
            "02_关键参数\n"
            "03_异常统计\n"
            "04_最终稳健中心\n"
            "05_最终稳健协方差\n\n"
            "01_FMCD结果已增加：\n"
            "异常数字化（背景=0，弱异常=1，中异常=2，强异常=3）\n\n"
            f"保存位置：\n{out_xlsx}"
        ),
    )

    open_output_folder(
        out_dir
    )


if __name__ == "__main__":
    try:
        main( )

    except SystemExit as exc:
        print(
            str(exc)
        )

    except Exception as exc:
        print(
            "\n"
            +
            "=" * 80
        )

        print(
            "程序运行失败"
        )

        print(
            "=" * 80
        )

        print(
            str(exc)
        )

        print(
            "\n完整错误位置："
        )

        traceback.print_exc()

        try:
            warn(
                "程序错误",
                (
                    f"程序运行出错：\n{exc}\n\n"
                    "详细错误位置已输出到控制台。"
                ),
            )

        except Exception:
            pass
