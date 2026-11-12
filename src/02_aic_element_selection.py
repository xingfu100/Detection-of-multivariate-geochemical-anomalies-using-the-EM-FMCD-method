import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.mixture import GaussianMixture
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import openpyxl
import itertools  # Generate combinations of arbitrary size.
import os  # Handle file paths.

# Configure Chinese font rendering.
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


def calculate_combo_aic(df_clean, selected_features, n_clusters, combo_num):
    """
    Calculate AIC for combinations of elements selected by individual AIC.

    Args:
        df_clean: Input data frame after cleaning.
        selected_features: Candidate element columns.
        n_clusters: Number of Gaussian mixture components.
        combo_num: Number of elements in each combination.

    Returns:
        A pair containing result records and invalid combination names.
    """
    combo_results = []
    invalid_combos = []

    # Generate distinct combinations from the selected elements.
    for combo in itertools.combinations(selected_features, combo_num):
        combo_name = "+".join(combo)
        try:
            # Retain rows with complete values for the selected elements.
            X_combo = df_clean[list(combo)].dropna().values
            if len(X_combo) < n_clusters:  # Skip combinations with fewer samples than clusters.
                invalid_combos.append(combo_name)
                print(f"  组合 {combo_name} 样本数不足，跳过")
                continue

            # Fit a Gaussian mixture model and calculate AIC.
            gmm = GaussianMixture(n_components=n_clusters,
                                  max_iter=1500,
                                  tol=1e-6,
                                  init_params='k-means++',
                                  covariance_type='diag',
                                  random_state=42)
            gmm.fit(X_combo)
            combo_aic = gmm.aic(X_combo)

            # Record the combination name and AIC.
            combo_results.append({
                '组合名称': combo_name,
                '组合AIC值': combo_aic
            })
            print(f"  组合 {combo_name}: 组合AIC={combo_aic:.4f}")

        except Exception as e:
            invalid_combos.append(combo_name)
            print(f"  组合 {combo_name} 计算失败: {e}")
            continue

    return combo_results, invalid_combos


def element_aic_analysis():
    root = tk.Tk()
    root.withdraw()

    # 1. Load the data.
    print("请选择数据文件...")
    file_path = filedialog.askopenfilename(filetypes=[('Excel', '*.xls;*.xlsx'), ('All', '*.*')])
    if not file_path:
        messagebox.showwarning('警告', '未选文件！')
        return
    try:
        df = pd.read_excel(file_path, header=0)
        # The first two columns are coordinates; subsequent columns are elements.
        all_features = df.columns.tolist()[2:]
        num_total_features = len(all_features)

        # Remove empty rows and rows with no element values.
        df_clean = df.dropna(how='all').copy()
        df_clean = df_clean.dropna(subset=all_features, how='all')

    except Exception as e:
        messagebox.showerror('错误', f'数据读取失败！\n错误信息: {e}')
        return

    # 2. Collect analysis parameters.
    # 2.1 Number of clusters.
    n_clusters_input = simpledialog.askstring('参数设定1/3', '请输入固定的聚类数 k (例如: 6):')
    if not n_clusters_input:
        return
    try:
        n_clusters = int(n_clusters_input)
        if n_clusters < 1:
            raise ValueError
    except:
        messagebox.showerror('错误', '输入的聚类数无效！必须是正整数')
        return

    # 2.2 Maximum combination size.
    max_combo_input = simpledialog.askstring('参数设定2/3',
                                             f'当前总元素数: {num_total_features}\n请输入最大组合数（例如输入4则计算2/3/4组合）:')
    if not max_combo_input:
        return
    try:
        max_combo = int(max_combo_input)
        if max_combo < 2:
            raise ValueError(f"组合数需≥2")
        # Allow a maximum larger than the number of available features.
        if max_combo > num_total_features:
            messagebox.showinfo('提示',
                                f'最大组合数({max_combo})超过总元素数({num_total_features})，超过部分将用所有有效元素计算！')
    except ValueError as ve:
        messagebox.showerror('错误', f'输入的最大组合数无效！\n原因: {ve}')
        return

    # 2.3 Additional elements to retain when screening by AIC.
    tolerance_input = simpledialog.askstring('参数设定3/3',
                                             f'容错数说明：\n- 两两组合(2) + 容错数(0) → 选AIC最低的2个元素做组合\n- 三三组合(3) + 容错数(2) → 选AIC最低的5个元素做组合\n请输入容错数（非负整数，例如0/2/10）:')
    if not tolerance_input:
        return
    try:
        tolerance_num = int(tolerance_input)
        if tolerance_num < 0:
            raise ValueError("容错数必须是非负整数")
        # Adjust the selected feature count at runtime.
    except ValueError as ve:
        messagebox.showerror('错误', f'输入的容错数无效！\n原因: {ve}')
        return

    print(
        f"\n===== 分析参数 =====\n聚类数: {n_clusters}\n最大组合数: {max_combo}\n容错数: {tolerance_num}\n====================")
    print("步骤1：计算所有元素的单个AIC值并排序（升序，AIC越低越好）...")

    # Step 1: calculate and rank single-element AIC values.
    element_aic = {}  # Map element names to AIC values.
    invalid_elements = []  # Record invalid elements.

    for feature in all_features:
        try:
            X_single = df_clean[[feature]].dropna().values
            if len(X_single) < n_clusters:
                invalid_elements.append(feature)
                continue

            gmm = GaussianMixture(n_components=n_clusters,
                                  max_iter=1500,
                                  tol=1e-6,
                                  init_params='k-means++',
                                  covariance_type='diag',
                                  random_state=42)
            gmm.fit(X_single)
            aic = gmm.aic(X_single)
            element_aic[feature] = aic
            print(f"  元素 {feature} 的AIC值: {aic:.4f}")

        except Exception as e:
            invalid_elements.append(feature)
            print(f"  元素 {feature} 计算AIC失败: {e}")

    # Check the number of valid elements.
    valid_features = list(element_aic.keys())
    num_valid_features = len(valid_features)
    if num_valid_features < 2:
        messagebox.showerror('错误', f'有效计算AIC的元素仅{num_valid_features}个，无法进行组合分析！')
        return

    # Sort elements by ascending AIC.
    sorted_elements = sorted(element_aic.items(), key=lambda x: x[1])
    sorted_element_names = [elem for elem, aic in sorted_elements]
    sorted_element_aic = [aic for elem, aic in sorted_elements]

    print(f"\n步骤1完成！有效元素数: {num_valid_features}")

    # Step 2: calculate AIC for each combination size.
    all_combo_results = {}  # Map combination sizes to result and invalid-combination lists.
    # Store the best result for each combination size.
    best_combo_results = []

    for combo_num in range(2, max_combo + 1):
        # Set the target to the combination size plus the tolerance.
        target_selected_num = combo_num + tolerance_num
        # Use all valid elements if fewer than the target are available.
        if target_selected_num > num_valid_features:
            selected_num = num_valid_features
            selected_features = sorted_element_names[:selected_num]
            print(
                f"\n步骤{combo_num}: 计算{combo_num}组合的AIC值（目标筛选{target_selected_num}个元素，但有效元素仅{num_valid_features}个，改用全部有效元素）...")
        else:
            selected_num = target_selected_num
            selected_features = sorted_element_names[:selected_num]
            print(f"\n步骤{combo_num}: 计算{combo_num}组合的AIC值（仅用AIC最低的{selected_num}个元素）...")

        print(f"  筛选的元素列表: {', '.join(selected_features)}")

        # Calculate AIC for combinations of this size.
        combo_res, invalid_res = calculate_combo_aic(
            df_clean=df_clean,
            selected_features=selected_features,
            n_clusters=n_clusters,
            combo_num=combo_num
        )
        all_combo_results[combo_num] = (combo_res, invalid_res)

        # Select the combination with the lowest AIC.
        if combo_res:
            best_combo = min(combo_res, key=lambda x: x['组合AIC值'])
            best_combo_results.append({
                '组合数': combo_num,
                '最优组合名称': best_combo['组合名称'],
                '最优AIC值': best_combo['组合AIC值'],
                '筛选元素数': selected_num
            })
            print(f"  ✨ {combo_num}组合最优结果: {best_combo['组合名称']} (AIC={best_combo['组合AIC值']:.4f})")
        else:
            best_combo_results.append({
                '组合数': combo_num,
                '最优组合名称': '无有效结果',
                '最优AIC值': None,
                '筛选元素数': selected_num
            })
            print(f"  ❌ {combo_num}组合无有效结果")

        # Check whether any valid results exist for this size.
        if not combo_res:
            messagebox.showwarning('警告', f'{combo_num}组合无有效计算结果！')

    # Step 3: save results and best combinations to one workbook.
    print("\n所有计算完成，选择保存路径...")
    save_path = filedialog.asksaveasfilename(
        defaultextension='.xlsx',
        filetypes=[('Excel', '*.xlsx')],
        initialfile=f"{os.path.splitext(os.path.basename(file_path))[0]}_低AIC元素组合分析.xlsx"
    )
    if not save_path:
        messagebox.showwarning('警告', '未选择保存路径，结果未保存！')
        return

    # Create the output workbook.
    wb = openpyxl.Workbook()

    # Sheet 1: ranking of single-element AIC values.
    ws1 = wb.active
    ws1.title = "单元素AIC排序"
    ws1.append(['元素名称', 'AIC值', '是否入选筛选池'])
    # Determine the maximum number of selected elements for marking.
    max_actual_selected_num = min(max_combo + tolerance_num, num_valid_features)
    for i, (elem, aic) in enumerate(sorted_elements):
        is_selected = "是" if i < max_actual_selected_num else "否"
        ws1.append([elem, round(aic, 4), is_selected])
    # Mark invalid elements.
    if invalid_elements:
        ws1.append(['', '', ''])
        ws1.append(['无效元素', '原因', ''])
        for elem in invalid_elements:
            ws1.append([elem, '样本不足/计算失败', ''])
    ws1.column_dimensions['A'].width = 15
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 15

    # Sheet 2: best combination for each combination size.
    ws_best = wb.create_sheet(title="各组合数最优AIC汇总")
    ws_best.append(['组合数', '筛选元素数', '最优组合名称', '最优AIC值'])
    for res in best_combo_results:
        aic_val = round(res['最优AIC值'], 4) if res['最优AIC值'] is not None else '无'
        ws_best.append([res['组合数'], res['筛选元素数'], res['最优组合名称'], aic_val])
    ws_best.column_dimensions['A'].width = 10
    ws_best.column_dimensions['B'].width = 15
    ws_best.column_dimensions['C'].width = 40
    ws_best.column_dimensions['D'].width = 20

    # Create a sheet for each combination size.
    for combo_num in range(2, max_combo + 1):
        combo_res, invalid_res = all_combo_results[combo_num]
        sheet_name = f"{combo_num}组合AIC"

        # Create the worksheet.
        ws = wb.create_sheet(title=sheet_name)
        # Write the combination and AIC column headings.
        ws.append(['组合名称', '组合AIC值'])

        # Write valid combinations in ascending AIC order.
        if combo_res:
            combo_res_sorted = sorted(combo_res, key=lambda x: x['组合AIC值'])
            for res in combo_res_sorted:
                ws.append([res['组合名称'], round(res['组合AIC值'], 4)])
        else:
            ws.append(['无有效组合结果', ''])

        # Mark invalid combinations.
        if invalid_res:
            # Separate sections with a blank row.
            ws.append(['', ''])
            ws.append([f'无效{combo_num}组合', '原因'])
            for combo in invalid_res:
                ws.append([combo, '样本不足/计算失败'])

        # Adjust column widths.
        ws.column_dimensions['A'].width = 10 * combo_num if combo_num <= 10 else 80  # Limit maximum column width.
        ws.column_dimensions['B'].width = 20

    # Save the workbook.
    wb.save(save_path)

    # Report the final summary.
    valid_combo_count = {k: len(v[0]) for k, v in all_combo_results.items()}
    msg = f'分析全部完成！\n\n'
    msg += f'核心参数：\n- 聚类数: {n_clusters}\n- 最大组合数: {max_combo}\n- 容错数: {tolerance_num}\n\n'
    msg += f'有效元素数: {num_valid_features}\n'
    msg += f'最大筛选元素数: {max_actual_selected_num}\n'
    for k in valid_combo_count:
        msg += f'有效{k}组合数: {valid_combo_count[k]}\n'
    msg += f'\n关键新增：\n✅ 各组合数最优AIC已汇总至「各组合数最优AIC汇总」Sheet\n'
    msg += f'\n结果已保存至：\n{save_path}'
    messagebox.showinfo('完成', msg)


if __name__ == '__main__':
    # Enter the cluster count, maximum combination size and tolerance in the interface.
    element_aic_analysis()