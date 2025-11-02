# Yeast-based FBP Caffeic Acid Quantification

## How to Use
Download the Git Hub directory.

```sh
git clone https://github.com/klocke3/Yeast-based-FBP.git
```

Prepare your .xlsx data file as a template. In each column, insert the RLU values from your samples in triplicate. At folder Yeast-based-FBP, save your data file and run the script using the following command:

```sh
cd /Yeast-based-FBP
```

```sh
python script.py -i data.xlsx -u micromolar
```

| Argument | Description                                    |
| -------- | ---------------------------------------------- |
| `-i`     | Input `.xlsx` file containing the raw RLU data |
| `-u`     | Unit of caffeic acid concentration  `micromolar` or  `millimolar`|

## Data Organization

Each sample should be placed in a separate sheet in the Excel file.

An example template is provided in the Example/ folder.


## Output

Plots: Saved in the `Results/graphics/` folder (one plot per sheet).

Summary table: Saved as `Result/summary_results.xlsx`, including:

Caffeic acid concentration

R² (correlation coefficient)

Average AUC (Area Under the Curve) and standard deviation (SD)

Slope and standard error (SE)

Intercetp and standard error (SE)

F value

Prob > F




