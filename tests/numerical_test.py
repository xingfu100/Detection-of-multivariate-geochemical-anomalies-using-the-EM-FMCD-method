"""Run deterministic checks of several numerical building blocks without a GUI."""
import importlib.util
from pathlib import Path
import numpy as np
import pandas as pd

source_dir = Path(__file__).resolve().parents[1] / "src"
def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, source_dir / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

preprocessing = load("preprocessing", "01_preprocessing.py")
aic_selection = load("aic_selection", "02_aic_element_selection.py")
em_clustering = load("em_clustering", "03_em_clustering.py")
robust_pca = load("robust_pca", "04_robust_pca.py")
fmcd = load("fmcd", "05_fmcd.py")

x = pd.DataFrame([[1., 2., 4.], [3., 3., 3.], [10., 20., 40.]], columns=["A", "B", "C"])
clr = preprocessing.clr_transform(x)
np.testing.assert_allclose(clr.sum(axis=1), 0, atol=1e-12)
np.testing.assert_allclose(clr.iloc[0], clr.iloc[2], atol=1e-12)
ilr, basis = preprocessing.ilr_transform(x)
v = basis.iloc[:, 1:].to_numpy(dtype=float)
np.testing.assert_allclose(v.T @ v, np.eye(2), atol=1e-12)
np.testing.assert_allclose(ilr.to_numpy(), clr.to_numpy() @ v, atol=1e-12)
h = robust_pca.构造Pivot_ILR正交基(3)
np.testing.assert_allclose(h.T @ h, np.eye(2), atol=1e-12)
np.testing.assert_allclose(h.sum(axis=0), 0, atol=1e-12)
assert em_clustering.gmm_parameter_count(2, 3, "diag") == 13
np.testing.assert_allclose(em_clustering.posterior_entropy(np.array([[0.5, 0.5]])), [np.log(2)])
np.testing.assert_allclose(fmcd.mahalanobis2(np.array([[1., 2.], [0., 0.]]), np.zeros(2), np.eye(2)), [5., 0.])
results, invalid = aic_selection.calculate_combo_aic(x, ["A", "B", "C"], 1, 2)
assert len(results) == 3 and not invalid
assert all(np.isfinite(row["组合AIC值"]) for row in results)
# Exercise robust PCA and FMCD on a deterministic, strictly positive composition.
rng = np.random.default_rng(47)
composition = np.exp(rng.normal(size=(100, 4)))
pca = robust_pca.运行ILR_MCD_RPCA(composition, ["A", "B", "C", "D"], 0.5, 42)
assert pca["scores_all"].shape == (100, 3)
np.testing.assert_allclose(pca["clr_loadings_all"].sum(axis=0), 0, atol=1e-10)
assert pca["clr_reconstruction_error"] < 1e-10
result = fmcd.fast_mcd(rng.normal(size=(100, 3)), starts=20, keep_best=5, seed=42)
assert result["final_md2"].shape == (100,)
assert np.all(np.isfinite(result["final_md2"])) and np.all(result["final_md2"] >= 0)
print("PASS: CLR, ILR, GMM, AIC, robust PCA and FMCD synthetic checks")
