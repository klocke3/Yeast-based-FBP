# Versão atualizada do script com:
# - Cálculo automático de LOD e LOQ
# - Cálculo da concentração por adição de padrão
# - Desvio (incerteza) da concentração
# - Substituição de "x when y=0" por "[caffeic acid] (concentração)"

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress, f
import os
import argparse

# 1️⃣ Parser
parser = argparse.ArgumentParser(description="Process Excel data and plot AUC with regression and ANOVA analysis")
parser.add_argument("-i", "--input", required=True, help="Input Excel file")
parser.add_argument("-u", "--unit", choices=["micromolar", "millimolar"], default="micromolar", help="Unit for caffeic acid concentration (micromolar or millimolar)")
args = parser.parse_args()

# 2️⃣ Map user input to chemical symbol
unit_map = {
    "micromolar": "µM",
    "millimolar": "mM"
}
unit_symbol = unit_map[args.unit]

excel_file = args.input
print(f"📂 Input file: {excel_file}")
print(f"⚗️ Unit for caffeic acid: {unit_symbol}")

# 3️⃣ Get sheet names
sheets = pd.ExcelFile(excel_file).sheet_names

# Create folder for saving graphs
output_folder = 'Results/graphs'
os.makedirs(output_folder, exist_ok=True)

summary_data = []

# 4️⃣ Loop through sheets
for sheet_name in sheets:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # Extract time vector
    time = df.iloc[3:33, 0].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce').reset_index(drop=True)

    # Group columns (3 replicates + 1 skip)
    cols = list(range(1, df.shape[1]))
    grouped_cols = []
    for i in range(0, len(cols), 4):
        group = tuple(cols[i:i+3])
        if len(group) == 3:
            grouped_cols.append(group)

    mean_areas, std_areas, labels = [], [], []

    for group in grouped_cols:
        c1, c2, c3 = group
        sub = df.iloc[3:33, [c1, c2, c3]].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce')
        areas = [np.trapz(sub.iloc[:, i], time) for i in range(sub.shape[1])]
        mean_areas.append(np.mean(areas))
        std_areas.append(np.std(areas, ddof=1))
        cell_name = str(df.iloc[0, c1])
        number = ''.join(filter(lambda x: x.isdigit() or x == '.', cell_name))
        labels.append(float(number) if number else np.nan)

    # Linear regression
    x = np.array(labels)
    y = np.array(mean_areas)
    y_err = np.array(std_areas)

    slope, intercept, r_value, p_value, std_err = linregress(x, y)
    y_fit = slope * x + intercept

    # ANOVA
    n = len(x)
    y_mean = np.mean(y)
    ss_reg = np.sum((y_fit - y_mean)**2)
    ss_res = np.sum((y - y_fit)**2)
    df_reg = 1
    df_res = n - 2
    ms_reg = ss_reg / df_reg
    ms_res = ss_res / df_res
    F_value = ms_reg / ms_res if ms_res != 0 else np.nan
    p_anova = 1 - f.cdf(F_value, df_reg, df_res) if not np.isnan(F_value) else np.nan

    # Standard errors
    s_yx = np.sqrt(ms_res)
    s_xx = np.sum((x - np.mean(x))**2)
    se_slope = s_yx / np.sqrt(s_xx)
    se_intercept = s_yx * np.sqrt(1/n + np.mean(x)**2 / s_xx)

    # Plot
    plt.figure(figsize=(8, 5))
    plt.errorbar(x, y, yerr=y_err, fmt='o', capsize=5, markersize=8, color='royalblue')
    plt.plot(x, y_fit, color='red', linestyle='--', label=f'Linear regression (r={r_value:.4f})')
    plt.xlabel(f'[Caffeic acid] samples ({unit_symbol})')
    plt.ylabel('AUC (RLU)')
    plt.title(f'Average RLU (± SD) — {sheet_name}')
    plt.ylim(bottom=0)
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    plt.legend()
    plt.tight_layout()
    filename = os.path.join(output_folder, f"{sheet_name.replace('/', '_')}.png")
    plt.savefig(filename, dpi=300)
    plt.close()

    # Concentration by standard addition
    if slope != 0:
        conc = -intercept / slope
    else:
        conc = np.nan

    # Uncertainty of concentration
    if slope != 0:
        conc_err = np.sqrt(
            (se_intercept / slope)**2 +
            ((intercept * se_slope) / (slope**2))**2
        )
    else:
        conc_err = np.nan

    # LOD and LOQ (using SE(intercept))
    LOD = (3.3 * se_intercept) / slope if slope != 0 else np.nan
    LOQ = (10 * se_intercept) / slope if slope != 0 else np.nan

    # Store summary
    summary_data.append({
        'Sheet name': sheet_name,
        f'[Caffeic acid] ({unit_symbol})': ', '.join([f'{val:.2f}' for val in labels]),
        'Average AUC': ', '.join([f'{val:.2f}' for val in mean_areas]),
        'SD (AUC)': ', '.join([f'{val:.2f}' for val in std_areas]),
        'R': f'{r_value:.3f}',
        'Slope': f'{slope:.3f}',
        'SE(Slope)': f'{se_slope:.2e}',
        'Intercept': f'{intercept:.3f}',
        'SE(Intercept)': f'{se_intercept:.2e}',
        'F value': f'{F_value:.3f}',
        'Prob > F': f'{p_anova:.3e}',
        'Concentration (± error)': f'{(abs(conc) / 0.99):.3f} ± {(conc_err /0.99) :.3f}',
        'LOD': f'{LOD:.3f}',
        'LOQ': f'{LOQ:.3f}'
    })

# Save table
path_save = 'Results'
os.makedirs(path_save, exist_ok=True)
summary_df = pd.DataFrame(summary_data)
summary_file = os.path.join(path_save, 'summary_results.xlsx')
summary_df.to_excel(summary_file, index=False)

print(f"\n✅ Summary table saved as: {summary_file}")
print(f"📊 Graphs saved in folder: '{output_folder}'")
print("🎯 Done! Includes concentration, uncertainty, LOD, LOQ, and corrected naming.")













