# Synthetic EM–RPCA–FMCD workflow

This example runs the EM clustering, robust PCA and FMCD analysis on 1,200 synthetic samples with 39 positive compositional variables. The coordinates are grid indices, not geographical coordinates. The example was constructed to illustrate differing geochemical backgrounds and an imposed anomaly; it is not an independent validation or a reproduction of the manuscript's numerical results. Real study data are not included.

## Installation and execution

Use Python 3 with Tkinter and install `numpy`, `pandas`, `scipy`, `scikit-learn`, `matplotlib` and `openpyxl`. For example:

```text
python -m pip install numpy pandas scipy scikit-learn matplotlib openpyxl
python examples/generate_synthetic.py
python examples/synthetic_workflow.py
python examples/plot_results.py
```

The random seed is 42. Running the workflow overwrites the files under `results`; copy them elsewhere before experimenting. The generator uses no study data. Results can vary slightly with library versions and numerical platforms.

## Files and manual reproduction

Each table is an individual CSV file; RPCA inputs are also provided as Excel workbooks. The first two columns of each input are artificial X and Y indices. Do not add an identifier column to the inputs: the desktop programs interpret later columns as variables. The row order is preserved. `03_em/assignments.csv` includes a one-based sample identifier for joining to the original dataset; cluster labels in this file are numbered 1–5, while `sample_comparison.csv` uses zero-based labels. The three programs retain the filenames and content of the supplied original scripts. The supplied Excel input files are snapshots: the automatic workflow regenerates CSV files only. After changing the input data, open each updated RPCA CSV and save it as XLSX before using the desktop RPCA program.

1. `results/workflow/01_raw/concentrations.csv`: simulated positive concentrations. There are no missing or zero values to replace. For every row, compute the natural logarithm of each of the 39 elements and subtract the mean of those 39 logarithms to obtain `02_clr/clr_all.csv`.
2. `03_em/input_clr.csv`: select the 14 indicated CLR columns from the full CLR table. Fit a Gaussian mixture with five components and seed 42 using `src/1_EM&AIC.py`; assignments are in `03_em/assignments.csv` and model trials in `trials.csv`. Preserve the row order when joining labels back to both raw and CLR tables.
3. `04_rpca/regional/input_raw.xlsx` and `04_rpca/cluster_01` through `cluster_05/input_raw.xlsx`: use the raw positive concentrations of the entire area or of the corresponding EM cluster with `src/2_RPCA.py`; the program applies its pivot ILR, fits MCD robust PCA with support fraction 0.5 and seed 42, and maps loadings back to CLR space. ILR basis, transformed data, scores, explained variance, loadings and the element-selection table are saved beside each input.
4. `05_fmcd/regional/input_clr.csv` and the five cluster-specific files: these use the selected element columns **from the full 39-element CLR**, subset to the corresponding rows. Run `src/3_FMCD.py` with seed 42. Per-model files provide the robust center and covariance, squared Mahalanobis distance, normalized index and anomaly level. The threshold uses the chi-square 0.975 quantile with degrees of freedom equal to the number of selected elements; index levels use 1, 1.5 and 2 as cut points.
5. `results/sample_comparison.csv`, `model_summary.csv` and `detection_comparison.csv` combine results. The figure scripts produce EM, loading and anomaly figures. The separate `examples/synthetic_ground_truth.csv` records the deliberately imposed simulation labels and is used only to score the synthetic scenario.

The example's selection rule is implemented in `select()` in `examples/synthetic_workflow.py`: it uses PC1–PC2 directions and lengths relative to Pb to propose three to six elements. It is an example rule; geological interpretation and manual selection may differ. The whole-region and cluster-specific associations can differ, so the FMCD comparison is not a controlled estimate of the isolated effect of clustering. No result from this synthetic example proves superiority on field data.

The interactive scripts provide Chinese dialogs and may export more intermediate calculations; the automatic script uses their computational functions directly. Re-running an interactive analysis requires the same parameters and selected variables shown in `element_selection.csv`.

## Study data and license

The original study dataset and true coordinates are not included. The synthetic example tests the workflow but cannot reproduce the manuscript's original numerical results. The authors should supply the actual data availability and access conditions in the manuscript when known. The project is distributed under the MIT license; see the root repository LICENSE.
