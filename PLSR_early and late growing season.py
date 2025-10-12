# Analysis using Principal Component Analysis (PCA) and Partial Least Squares (PLS)
# for early and late season variables.
import warnings
from sklearn.decomposition import PCA
from sklearn.cross_decomposition import PLSRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import r2_score
import pandas as pd
import numpy as np
import re

# Suppress FutureWarnings from scikit-learn
warnings.filterwarnings("ignore", category=FutureWarning, module="sklearn")


# Function to check if a string contains Chinese characters and is not a time-scale indicator.
def is_plant_column(s):
    # Returns True if the string contains a Chinese character and does not contain time markers.
    return re.search("[\u4e00-\u9fff]", s) and not any(time_scale in s for time_scale in ['4-6', '7-9', '4-9'])


# Load the data
file_path = 'C:/Users/lenovo/Desktop/土壤站点数据/11-14/数据-1/数据-1/逐物种结果提取/56065xxc.xlsx'  # Change to your file path
data = pd.read_excel(file_path, sheet_name='data')

# Get the column names for the plant species (LSD columns)
plant_columns = [col for col in data.columns if is_plant_column(col)]

# Define the time periods for the analysis
time_periods = ['4-6', '7-9']

# Initialize a DataFrame to store the results
results_df = pd.DataFrame()

# Initialize the StandardScaler
scaler = StandardScaler()

# Iterate over each plant LSD column (i.e., each species)
for plant_column in plant_columns:
    # Define the corresponding GUD (Growing Up Day/萌芽期) column name
    gud_column = f'{plant_column}gud'

    # If the GUD column for this species does not exist, skip to the next species
    if gud_column not in data.columns:
        continue

    # Construct the list of independent variables, including environmental factors and the GUD column
    variables = [f'0-40cm{period}' for period in time_periods] + \
                [f'pre{period}' for period in time_periods] + \
                [f'TMEAN{period}' for period in time_periods] + \
                [f'SSD{period}' for period in time_periods] + \
                [gud_column]

    # Set the dependent variable (y) as the plant's LSD column and drop missing values
    y = data[plant_column].dropna()

    # Select the independent variables (X) and the dependent variable (y)
    # X includes GUD and other environmental variables
    X = data.loc[y.index, variables].dropna()
    # Update y to align with the filtered (non-missing) rows of X
    y = y.loc[X.index]

    # Check if X or y is empty or if there are not enough samples for the analysis
    if X.empty or len(y) < 2:
        continue  # If so, skip this species

    # Standardize the independent variables
    X_scaled = scaler.fit_transform(X)

    # Use PCA to calculate the variance explained by components
    pca = PCA()
    pca.fit(X_scaled)

    # Calculate the variance explained by each principal component
    explained_variance_ratio = pca.explained_variance_ratio_
    cumulative_variance_ratio = np.cumsum(explained_variance_ratio)

    # Print the variance explained ratio for each component
    print(f"Explained variance ratio for each component: {explained_variance_ratio}")
    print(f"Cumulative explained variance ratio: {cumulative_variance_ratio}")

    # Determine the number of components to use (e.g., reaching 90% cumulative variance)
    num_components = np.argmax(cumulative_variance_ratio >= 0.90) + 1
    print(f"Number of components selected: {num_components}")

    # Build and fit the PLS (Partial Least Squares) regression model
    pls = PLSRegression(n_components=num_components)
    pls.fit(X_scaled, y)

    # Calculate VIP (Variable Importance in Projection) scores
    t = pls.x_scores_
    w = pls.x_weights_
    q = pls.y_loadings_
    p, h = w.shape
    vips = np.zeros((p,))
    s = np.diag(t.T @ t @ q.T @ q).reshape(h, -1)
    total_s = np.sum(s)
    for i in range(p):
        weight = np.array([(w[i, j] / np.linalg.norm(w[:, j])) ** 2 for j in range(h)])
        vips[i] = np.sqrt(p * (s.T @ weight) / total_s)

    # Append the results for each variable to the main DataFrame
    for i, variable in enumerate(variables):
        result = {
            'Plant': plant_column,
            'GUD_Variable': gud_column,
            'Variable': variable,
            'Coefficient': pls.coef_.ravel()[i],
            'VIP': vips[i],
            'R2': r2_score(y, pls.predict(X_scaled)) if i == 0 else np.nan  # Record R2 value only once per model
        }
        results_df = pd.concat([results_df, pd.DataFrame([result])], ignore_index=True)

# Export the final results to an Excel file
output_file_path = 'C:/Users/lenovo/Desktop/计算贡献25_1/5变量(TMEAN)/控制主成分的前后期/56065xxc.xlsx'  # Change to your desired output file path
results_df.to_excel(output_file_path, index=False)

print(f"Results exported to {output_file_path}")