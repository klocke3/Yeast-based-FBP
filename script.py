import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress, f
import os
import argparse

# 1️⃣ Parser
parser = argparse.ArgumentParser(description="Process Excel data and plot AUC with regression and lack of fit analysis")

parser.add_argument(
    "-i", "--input",
    required=True,
    help="Input Excel file"
)

parser.add_argument(
    "-u", "--unit",
    choices=["micromolar", "millimolar"],
    default="micromolar",
    help="Unit for caffeic acid concentration (micromolar or millimolar)"
)

args = parser.parse_args()

# 2️⃣ Map user input para símbolo químico
unit_map = {
    "micromolar": "µM",
    "millimolar": "mM"
}

unit_symbol = unit_map[args.unit]

print(f"📂 Input file: {args.input}")
print(f"⚗️ Unit for caffeic acid: {unit_symbol}")

excel_file = args.input

# 3️⃣ Obter todas as planilhas
sheets = pd.ExcelFile(excel_file).sheet_names

# Criar pasta para salvar gráficos
output_folder = 'Results/graphs'
os.makedirs(output_folder, exist_ok=True)

# Lista para armazenar resultados
summary_data = []

# 4️⃣ Loop em todas as planilhas
for sheet_name in sheets:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # Extrair vetor de tempo (coluna A, linhas 4–33)
    time = df.iloc[3:33, 0].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce').reset_index(drop=True)

    # Criar grupos de colunas (3 colunas + 1 pulada)
    cols = list(range(1, df.shape[1]))
    grouped_cols = []
    for i in range(0, len(cols), 4):
        group = tuple(cols[i:i+3])
        if len(group) == 3:
            grouped_cols.append(group)

    mean_areas = []
    std_areas = []
    labels = []

    for group in grouped_cols:
        c1, c2, c3 = group
        sub = df.iloc[3:33, [c1, c2, c3]].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce')

        # Integrar (AUC) para cada réplica
        areas = [np.trapz(sub.iloc[:, i], time) for i in range(sub.shape[1])]

        # Média e desvio padrão
        mean_areas.append(np.mean(areas))
        std_areas.append(np.std(areas, ddof=1))

        # Extrair concentração numérica da primeira célula do grupo
        cell_name = str(df.iloc[0, c1])
        number = ''.join(filter(lambda x: x.isdigit() or x == '.', cell_name))
        labels.append(float(number) if number else np.nan)

    # Regressão linear
    x = np.array(labels)
    y = np.array(mean_areas)
    y_err = np.array(std_areas)

    slope, intercept, r_value, p_value, std_err = linregress(x, y)
    y_fit = slope * x + intercept

    # ---------- Teste de Lack of Fit ----------
    all_x, all_y = [], []
    for group, conc in zip(grouped_cols, labels):
        c1, c2, c3 = group
        sub = df.iloc[3:33, [c1, c2, c3]].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce')
        areas = [np.trapz(sub.iloc[:, i], time) for i in range(sub.shape[1])]
        for val in areas:
            all_x.append(conc)
            all_y.append(val)

    all_x = np.array(all_x)
    all_y = np.array(all_y)
    y_pred_all = slope * all_x + intercept

    ss_res = np.sum((all_y - y_pred_all) ** 2)

    # Soma dos quadrados do erro puro (dentro das replicatas)
    ss_pe = 0
    n_rep = 3
    for conc in np.unique(all_x):
        mask = all_x == conc
        y_rep = all_y[mask]
        ss_pe += np.sum((y_rep - np.mean(y_rep)) ** 2)
    df_pe = len(np.unique(all_x)) * (n_rep - 1)

    # Soma dos quadrados da falta de ajuste
    ss_lof = ss_res - ss_pe
    df_lof = len(np.unique(all_x)) - 2  # dois parâmetros (slope e intercept)

    ms_lof = ss_lof / df_lof
    ms_pe = ss_pe / df_pe

    F_lof = ms_lof / ms_pe if ms_pe != 0 else np.nan
    p_lof = 1 - f.cdf(F_lof, df_lof, df_pe) if not np.isnan(F_lof) else np.nan
    # -----------------------------------------

    # Plotagem
    plt.figure(figsize=(8, 5))
    plt.errorbar(x, y, yerr=y_err, fmt='o', capsize=5, markersize=8, color='royalblue')
    plt.plot(x, y_fit, color='red', linestyle='--', label=f'Linear regression (r={r_value:.2f})')
    plt.xlabel(f'[Caffeic acid] ({unit_symbol})')
    plt.ylabel('AUC (RLU)')
    plt.title(f'Average RLU (± SD) — {sheet_name}')
    plt.ylim(bottom=0)
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    plt.legend()
    plt.tight_layout()

    filename = os.path.join(output_folder, f"{sheet_name.replace('/', '_')}.png")
    plt.savefig(filename, dpi=300)
    plt.close()
    print(f"✅ Graph saved: {filename}")

    # Calcular x quando y=0
    x_zero = -intercept / slope if slope != 0 else np.nan

    # Armazenar resultados
    summary_data.append({
        'Sheet name': sheet_name,
        f'[Caffeic acid] ({unit_symbol})': ', '.join([f'{val:.2f}' for val in labels]),
        'Average AUC': ', '.join([f'{val:.2f}' for val in mean_areas]),
        'SD (AUC)': ', '.join([f'{val:.2f}' for val in std_areas]),
        'R': f'{r_value:.3f}',
        'Slope': f'{slope:.3f}',
        'Intercept': f'{intercept:.3f}',
        'Regression p-value': f'{p_value:.3e}',
        'Lack of fit F': f'{F_lof:.3f}',
        'Lack of fit p-value': f'{p_lof:.3e}' if not np.isnan(p_lof) else 'N/A',
        'x when y=0': f'{x_zero:.3f}' if not np.isnan(x_zero) else 'N/A'
    })

# 5️⃣ Salvar resultados
path_save = '/Results'
os.makedirs(path_save, exist_ok=True)

summary_df = pd.DataFrame(summary_data)

summary_file = os.path.join(path_save, 'summary_results.xlsx')
summary_df.to_excel(summary_file, index=False)

print(f"\n✅ Summary table saved as: {summary_file}")
print(f"📊 Graphs saved in folder: '{output_folder}'")
print("🎯 Done! Includes regression significance and lack-of-fit analysis.")



