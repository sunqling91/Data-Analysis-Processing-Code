# Calculate the trend of LSD shifts in time series for each species
import pandas as pd
from scipy import stats


def calculate_trend(data):
    """
    Calculates the trend of the full yellowing and withering stage over the years for a single plant.

    Args:
      data: DataFrame, containing the full yellowing and withering stage data for a single plant across different years.

    Returns:
      tuple: A tuple containing slope, p-value, standard error, and std_resid. Returns (None, None, None, None) if there are fewer than 2 data points.
    """
    if len(data) < 2:
        return None, None, None, None

    # Perform linear regression to get slope, intercept, r-value, p-value, and standard error.
    slope, intercept, r_value, p_value, std_err = stats.linregress(data['年份'], data['黄枯盛期'])

    # Calculate the standard deviation of the residuals.
    predicted_values = slope * data['年份'] + intercept
    residuals = data['黄枯盛期'] - predicted_values
    std_resid = residuals.std()

    return slope, p_value, std_err, std_resid


def calculate_all_trends(excel_file, sheet_name='Sheet1', station_col='区站号', plant_col='植物中文名', year_col='年份',
                         value_col='黄枯盛期'):
    """
    Calculates the trend of the full yellowing and withering stage over the years for each plant in an Excel sheet.

    Args:
        excel_file (str): The path to the Excel file.
        sheet_name (str): The name of the sheet, defaults to 'Sheet1'.
        station_col (str): The column name for the station ID, defaults to '区站号'.
        plant_col (str): The column name for the plant's Chinese name, defaults to '植物中文名'.
        year_col (str): The column name for the year, defaults to '年份'.
        value_col (str): The column name for the full yellowing and withering stage, defaults to '黄枯盛-期'.

    Returns:
        DataFrame: A DataFrame containing the trend calculation results for each plant. Each row includes:
                   '区站号', '植物中文名', 'slope', 'p_value', 'std_err', 'std_resid'.
    """

    # Read the Excel file into a pandas DataFrame.
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    results = []

    # Group the DataFrame by both station ID and plant name.
    for (station, plant), group_data in df.groupby([station_col, plant_col]):
        # Calculate the trend for each group.
        slope, p_value, std_err, std_resid = calculate_trend(group_data)
        # Append the results to the list.
        results.append({
            station_col: station,
            plant_col: plant,
            'slope': slope,
            'p_value': p_value,
            'std_err': std_err,
            'std_resid': std_resid
        })

    # Convert the list of dictionaries to a DataFrame.
    return pd.DataFrame(results)


# Example usage (please replace with your actual file path)
excel_file = 'C:/Users/lenovo/Desktop/94-17.xlsx'  # Replace with your Excel file name
results_df = calculate_all_trends(excel_file)
print(results_df)

# If needed, save the results to a new Excel file:
results_df.to_excel('C:/Users/lenovo/Desktop/lsdresult.xlsx', index=False)
