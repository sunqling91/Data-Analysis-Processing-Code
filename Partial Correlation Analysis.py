# Partial correlation for the full period 1-23 (without controlling for GUD)
import pandas as pd
import pingouin as pg
from sklearn.preprocessing import StandardScaler

# Load the data
file_path = 'C:/Users/lenovo/Desktop/土壤站点数据/11-14/数据-1/数据-1/逐物种结果提取/52765ml.xlsx'  # Replace with your file path
sheet_name = '数据'  # The name of the sheet where the data is stored
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Automatically identify the species name
# It finds the column with 'gud' in its name and removes 'gud' to get the base species name.
species_name = [col.replace('gud', '').strip() for col in df.columns if 'gud' in col][0]
dependent_var = species_name  # The dependent variable column, e.g., "SpeciesName"
control_var = f"{species_name}gud"  # The GUD column, e.g., "SpeciesNamegud"

# Define the columns for analysis
analysis_columns = [
    'pre4-9',
    'TMEAN4-9',
    'SSD4-9',
    '0-40cm4-9',
    control_var,
    dependent_var
]

# Check if all required columns exist in the DataFrame
missing_cols = [col for col in analysis_columns if col not in df.columns]
if missing_cols:
    raise ValueError(f"The following columns are missing from the data: {missing_cols}")

# Extract and standardize the data
data = df[analysis_columns].dropna()  # Drop rows with missing values
scaler = StandardScaler()
scaled_data = scaler.fit_transform(data)
scaled_df = pd.DataFrame(scaled_data, columns=analysis_columns)

# Define the variable grouping rules
groups = {
    'pre': ['pre4-9'],  # ['pre4-6']、['pre7-9']
    'soil_moisture': ['0-40cm4-9'],  # ['0-40cm4-6']、['0-40cm7-9']
    'temperature': ['TMEAN4-9'],  # ['TMEAN4-6']、['TMEAN7-9']
    'radiation': ['SSD4-9'],  # ['SSD4-6']、['SSD7-9']
    'gud': [control_var]
}

# A list to store the partial correlation analysis results
results = []

# Perform partial correlation analysis: control variables according to grouping rules
for group_name, group_vars in groups.items():
    for var in group_vars:
        # Check if the variable is one of the main environmental factors
        if var in ['pre4-9', 'TMEAN4-9', 'SSD4-9', '0-40cm4-9']:
            # For these variables, do not control for the 'gud' column
            covars = [col for col in analysis_columns if col not in [var, dependent_var, control_var]]
        else:
            # For other cases (mainly the 'gud' variable itself), control for all other variables
            covars = [col for col in analysis_columns if col not in [var, dependent_var]]

        # Calculate the partial correlation
        partial_corr = pg.partial_corr(data=scaled_df, x=var, y=dependent_var, covar=covars, method='spearman')

        # Format and append the results to the list
        results.append({
            'Independent Variable': var,
            'Dependent Variable': dependent_var,
            'Partial Correlation Coefficient': partial_corr['r'].values[0],  # Extract the single numerical value
            'p-value': partial_corr['p-val'].values[0],  # Extract the single numerical value
            'Control Variables': ', '.join(covars),
            'Variable Group': group_name
        })

# Convert the results list to a pandas DataFrame
results_df = pd.DataFrame(results)

# Save the results to the specified path
output_file = 'C:/Users/lenovo/Desktop/output.xlsx'
results_df.to_excel(output_file, index=False)
print(f"Analysis results have been saved to {output_file}")
