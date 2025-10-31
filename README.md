# Yeast-based-FBP

How to Use

Prepare your .xlsx data file as a template. In each column, insert the RLU values from your samples in triplicate.

Run the script using the following command:

python script.py -i data.xlsx -u micromolar


The -i argument specifies the input .xlsx file.

The -u argument defines the unit of caffeic acid concentration.
Accepted values are:

"micromolar" → displays units as µM

"millimolar" → displays units as mM

The generated plots will be saved in the Graphics/ folder, and the calculated results (including caffeic acid concentration, AUC, R², and p-value) will be stored in summary_results.xlsx.

For each sample, use a separate sheet name in the Excel file.
An example data template is available in the Example/ folder.
