# EM–FMCD geochemical anomaly analysis

Python scripts supporting an EM–FMCD workflow for multivariate geochemical anomaly analysis. They provide compositional preprocessing, Gaussian mixture clustering, robust principal component analysis (RPCA), and fast minimum covariance determinant (FMCD) distance estimation. The scripts are independent desktop applications with Chinese-language dialogs and outputs; source-code comments and docstrings are in English.

## Workflow and source files

The numbering follows the research procedure described in the manuscript:

1. **Data preparation** — `src/01_preprocessing.py` cleans concentration data, replaces nonpositive and missing values and can perform CLR, ILR, log10, and standardization. For the manuscript's EM and FMCD steps, produce CLR data. The first two columns are coordinates and remaining columns contain elements. Detection limits are hard-coded for the original study: verify their units and values against your own laboratory data before use. The script does not convert concentration units.
2. **Clustering indicators** — `src/02_aic_element_selection.py` is an auxiliary AIC-based element-combination screening tool. The manuscript also uses geological knowledge and correlation screening to choose indicators; the script does not replace those decisions. The user chooses the number of components, maximum combination size, and screening tolerance.
3. **Population separation** — `src/03_em_clustering.py` fits Gaussian mixtures using EM, compares AIC across cluster counts and seeds, and outputs cluster assignments. Select the cluster count by considering AIC, cluster sizes, and geological interpretation, rather than AIC alone.
4. **Element associations** — `src/04_robust_pca.py` carries out pivot ILR and MCD-based robust PCA and maps loadings back to CLR element space. Run separately for each population using strictly positive, unit-consistent concentration data corresponding to that population. Do **not** feed CLR or Z-score data into this script. Select target-related associations using loadings and geological interpretation.
5. **Anomaly detection** — `src/05_fmcd.py` estimates robust squared Mahalanobis distances and anomaly levels using selected, already preprocessed variables. For the manuscript workflow, use the appropriate selected CLR element columns for each population. The script does not perform CLR or other preprocessing. A whole-region run can also be used for comparison.

These are interactive tools, not an automatic pipeline: prepare the appropriate input table and select the variables for each stage. The repository does not include the study's original dataset, intermediate files, geological maps, or all manual geological decisions; the quick tests therefore do not reproduce the paper's numerical results. Excel sheet names, input/output labels and dialogs remain in Chinese for compatibility with the original workflow.

## Installation and use

Install Python 3 with Tkinter and a graphical desktop. From a terminal in the repository root:

```text
python -m venv .venv
```

On Windows, activate with `.venv\Scripts\activate`; on macOS/Linux use `source .venv/bin/activate`. Then install dependencies and launch a script:

```text
python -m pip install -r requirements.txt
python src/01_preprocessing.py
```

Run the other numbered scripts similarly. Follow each dialog to choose an input file, worksheet where applicable, variables, parameters, and an output directory. Most stages expect the first two columns to be coordinates; review the dialog and script description before running. Input support varies between programs; `openpyxl` and `xlrd` support Excel formats where the scripts accept them. Preserve sample identifiers externally when joining cluster membership to raw concentrations and CLR data. Ensure that each cluster has sufficient observations for the chosen number of variables and review convergence, covariance rank and model diagnostics.

## Quick tests

From the repository root, first check Python syntax:

```text
python tests/quick_test.py
```

After installing the dependencies, run synthetic-data checks:

```text
python tests/numerical_test.py
```

The numerical test checks selected CLR and ILR properties, mixture-model quantities, Mahalanobis distance and element-combination AIC. It does not test all parameter choices, GUI interactions, full workflows or reproduce manuscript results. These checks were run on Windows with Python 3.14.5, NumPy 2.5.4, pandas 3.1.1, SciPy 1.18.1, scikit-learn 1.9.1, Matplotlib 3.11.2 and openpyxl 3.1.5.

## License

MIT; see [LICENSE](LICENSE).
