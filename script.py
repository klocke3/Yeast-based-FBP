import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
import os
import argparse

# 1️⃣ Parser
parser = argparse.ArgumentParser(description="Process Excel data and plot AUC with regression")

parser.add_argument(
    "-i", "--input",
    required=True,
    help="Input Excel file"
)

parser.add_argument(
    "-u", "--unit",
    choices=["micromolar", "millimolar"],  # o usuário digita isso
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

print(f"Input file: {args.input}")
print(f"Unit for caffeic acid: {unit_symbol}")

excel_file = args.input

# 2️⃣ Get all sheet names
sheets = pd.ExcelFile(excel_file).sheet_names

# Create folder to save graphs
output_folder = 'Yeast-based-FBP/Results/graphs'
os.makedirs(output_folder, exist_ok=True)

# Prepare list to store summary results
summary_data = []

# 3️⃣ Loop through all sheets
for sheet_name in sheets:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # Extract time vector (column A, rows 4–33)
    time = df.iloc[3:33, 0].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce').reset_index(drop=True)

    # Create list of column groups (3 columns + 1 skipped)
    cols = list(range(1, df.shape[1]))
    grouped_cols = []
    for i in range(0, len(cols), 4):
        group = tuple(cols[i:i+3])
        if len(group) == 3:
            grouped_cols.append(group)

    # Calculate AUC for each group
    mean_areas = []
    std_areas = []
    labels = []

    for group in grouped_cols:
        c1, c2, c3 = group
        sub = df.iloc[3:33, [c1, c2, c3]].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce')

        # Integrate (AUC) for each replicate
        areas = [np.trapz(sub.iloc[:, i], time) for i in range(sub.shape[1])]

        # Mean and SD
        mean_areas.append(np.mean(areas))
        std_areas.append(np.std(areas, ddof=1))

        # Extract numeric label from first cell of group (row 1)
        cell_name = str(df.iloc[0, c1])
        number = ''.join(filter(lambda x: x.isdigit() or x == '.', cell_name))
        labels.append(float(number) if number else np.nan)

    # Linear regression
    x = np.array(labels)
    y = np.array(mean_areas)
    y_err = np.array(std_areas)

    slope, intercept, r_value, p_value, std_err = linregress(x, y)
    y_fit = slope * x + intercept

    # Plot with error bars and regression line
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

    # Save plot as PNG
    filename = os.path.join(output_folder, f"{sheet_name.replace('/', '_')}.png")
    plt.savefig(filename, dpi=300)
    plt.close()
    print(f"✅ Graph saved: {filename}")

    # Compute x when y=0
    x_zero = -intercept / slope if slope != 0 else np.nan

    # Store summary results
    summary_data.append({
        'Sheet name': sheet_name,
        f'[Caffeic acid] ({unit_symbol})': ', '.join([f'{val:.2f}' for val in labels]),
        'Average AUC': ', '.join([f'{val:.2f}' for val in mean_areas]),
        'SD (AUC)': ', '.join([f'{val:.2f}' for val in std_areas]),
        'R': f'{r_value:.3f}',
        'Slope': f'{slope:.3f}',
        'Intercept': f'{intercept:.3f}',
        'p-value': f'{p_value:.3e}',
        'x when y=0': f'{x_zero:.3f}' if not np.isnan(x_zero) else 'N/A'
    })

# Criar pasta de resultados relativa ao diretório atual
path_save = 'Yeast-based-FBP/Results'
os.makedirs(path_save, exist_ok=True)

# Criar DataFrame de resumo
summary_df = pd.DataFrame(summary_data)

# Salvar o Excel na pasta criada
summary_file = os.path.join(path_save, 'summary_results.xlsx')
summary_df.to_excel(summary_file, index=False)

print(f"✅ Summary table saved as: {summary_file}")
print(f"📊 Graphs saved in folder: '{output_folder}'")

