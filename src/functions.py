import pandas as pd
import numpy as np
import re
from itertools import combinations
import string
import xlsxwriter
from scipy.stats import chi2_contingency
from statsmodels.formula.api import ols
from scipy.stats import gaussian_kde
from scipy.interpolate import interp1d
import warnings
warnings.filterwarnings(
    "ignore", 'This pattern is interpreted as a regular expression, and has match groups.')

# %load Functions.py


def load_tool_choices(filename_tool, label_colname, keep_cols=False):
    
    """
    Loads and processes the 'choices' sheet from a KoboToolbox tool Excel file.
    
    Parameters:
    ----------
    filename_tool : str
        The file path of the Excel file containing the tool.
        The file should contain a sheet named 'choices'.
        
    label_colname : str
        The name of the column in the 'choices' sheet that contains
        the labels or text you want to extract.
        
    keep_cols : bool, optional
        If True, keeps all columns from the 'choices' sheet.
        If False (default), only keeps the 'list_name', 'name',
        and the column specified by `label_colname`.

    Returns:
    -------
    pandas.DataFrame
        A DataFrame containing the cleaned and processed 'choices' data,
        filtered to include only the selected columns (if `keep_cols` is False).
    """
    
    tool_choices = pd.read_excel(
        filename_tool, sheet_name="choices", dtype="str")

    if not keep_cols:
        tool_choices = tool_choices[['list_name', 'name', label_colname]]

    # Remove rows with missing values in 'list_name' column
    tool_choices = tool_choices.dropna(subset=['list_name'])

    # Keep only distinct rows
    tool_choices = tool_choices.drop_duplicates()

    # Convert to DataFrame
    tool_choices = pd.DataFrame(tool_choices)

    return (tool_choices)


def load_tool_survey(filename_tool, label_colname, keep_cols=False):
    
    """
    Loads and processes the 'survey' sheet from a KoboToolbox tool Excel file.
    Extracts relevant columns based on language and type of questions, and 
    identifying the datasheet each question belongs to.

    Parameters:
    ----------
    filename_tool : str
        The file path of the Excel file containing the tool.
        The file should contain a sheet named 'survey'.
        
    label_colname : str
        The name of the column containing the labels in a specific language (e.g., 'label::English').
        Used to filter columns based on language.
        
    keep_cols : bool, optional
        If True, all columns from the 'survey' sheet will be retained.
        If False (default), only specific columns such as label, hint, and required message 
        (for the specified language) are kept.

    Returns:
    -------
    pandas.DataFrame
        A DataFrame containing the cleaned and processed 'survey' data, 
        including relevant question types (e.g., 'select_one', 'select_multiple', 'integer', 'decimal'), 
        filtered columns based on the specified language, and the datasheet each question belongs to.
    """
    
    tool_survey = pd.read_excel(
        filename_tool, sheet_name="survey", dtype="str")

    tool_survey = tool_survey.dropna(subset=['type'])

    tool_survey['q.type'] = tool_survey['type'].apply(
        lambda x: re.split(r'\s', x)[0])
    tool_survey['list_name'] = tool_survey['type'].apply(
        lambda x: re.split(r'\s', x)[1] if re.match(r'select_', x) else None)

    # Select only relevant columns
    if not keep_cols:
        lang_code = re.split(r'::', label_colname, maxsplit=1)[1]
        lang_code = re.sub(r'\(', r'\\(', lang_code)
        lang_code = re.sub(r'\)', r'\\)', lang_code)
        cols_to_keep = tool_survey.columns[(tool_survey.columns.str.contains(f'((label)|(hint)|(constraint_message)|(required_message))::{lang_code}')) |
                                           (~tool_survey.columns.str.contains(r'((label)|(hint)|(constraint_message)|(required_message))::'))]
        tool_survey = tool_survey[cols_to_keep]


    # Find which data sheet question belongs to
    tool_survey['datasheet'] = None
    sheet_name = "main"
    for i, toolrow in tool_survey.iterrows():
        if re.search(r'begin[_ ]repeat', toolrow['type']):
            sheet_name = toolrow['name']
        elif re.search(r'end[_ ]repeat', toolrow['type']):
            sheet_name = "main"
        elif not re.search(r'((end)|(begin))[_ ]group', toolrow['type'], re.IGNORECASE):
            tool_survey.loc[i, 'datasheet'] = sheet_name

    tool_survey = tool_survey[tool_survey['q.type'].isin(['select_one','select_multiple','integer','decimal'])]
    
    return tool_survey


def map_names(column_name, column_values_name, summary_table, tool_survey, tool_choices,label_col, na_include=False):
        
    """
    Maps the values in the `summary_table` column (specified by `column_values_name`) to corresponding labels
    found in `tool_choices`. The mapping is based on a survey structure found in `tool_survey`.

    Parameters:
    -------
        column_name (str): The name of the column in `summary_table` whose first value 
        matches a corresponding list in `tool_survey`.
        
        column_values_name (str): The name of the column in `summary_table` whose values
        will be mapped to corresponding labels.
        
        summary_table (pd.DataFrame): The data table containing the columns to be mapped.
        
        tool_survey (pd.DataFrame): The survey structure containing the mapping 
        of column names to list names.
        
        tool_choices (pd.DataFrame): The choices structure containing possible options 
        (list_name) and their corresponding labels.
        
        label_col (str): The name of the column in `tool_choices` that contains the labels for mapping.
        
        na_include (bool, optional): Whether to include an additional mapping for missing values 
        ('No data available (NA)'). Defaults to False.

    Returns:
    -------
        pd.DataFrame: The updated `summary_table` with the values in `column_values_name` mapped to corresponding labels.
    """
    
    # get the shortlist of choices per variable
    choices_shortlist = tool_choices[
        tool_choices['list_name'].values == tool_survey[tool_survey['name']
                                                        == summary_table[column_name][0]]['list_name'].values
    ][['name', label_col]]
    # build a dictionary
    mapping_dict = dict(
        zip(choices_shortlist['name'], choices_shortlist[label_col]))
    # add an NA entry to the dictionary if the parameter was chosen
    if na_include is True:
        mapping_dict['No_data_available_NA'] = 'No data available (NA)'
    # if something was missing from the tool, it will be its own label
    for value in summary_table[column_values_name]:
        if value not in mapping_dict:
            mapping_dict[value] = value
    # None breaks everything. Trying to change it
    mapping_dict['none'] = 'None '
    # add the _orig column with base names
    summary_table[column_values_name+'_orig']=summary_table[column_values_name].copy()
    # add the correct labels to the output table
    summary_table[column_values_name] = summary_table[column_values_name].map(
        mapping_dict)
    return summary_table


def map_names_ls(column_name, values_list, tool_survey, tool_choices,label_col, na_include=False):
    
    """
    Maps a list of values to corresponding labels from the 'choices' sheet in a KoboToolbox tool,
    based on a specified column in the 'survey' sheet. Is only used in the significance checks

    Parameters:
    ----------
    column_name : str
        The name of the column in the 'survey' sheet for which the values are being mapped.
        
    values_list : list
        A list of values that need to be mapped to their corresponding labels.

    tool_survey : pandas.DataFrame
        The DataFrame containing the 'survey' sheet data.
        
    tool_choices : pandas.DataFrame
        The DataFrame containing the 'choices' sheet data.
        
    label_col : str
        The name of the column in the 'choices' sheet that contains the labels 
        to map to the values in `values_list`.
        
    na_include : bool, optional
        If True, includes a mapping for missing data ('No_data_available_NA'). 
        Default is False.

    Returns:
    -------
    list
        A list of values from `values_list`, mapped to their corresponding labels 
        from the 'choices' sheet, with unmatched values preserved and 'none' explicitly mapped to 'None '.
    """
    # get the shortlist of choices per variable
    choices_shortlist = tool_choices[
        tool_choices['list_name'].values == tool_survey[tool_survey['name']== column_name]['list_name'].values
    ][['name', label_col]]
    # build a dictionary
    mapping_dict = dict(
        zip(choices_shortlist['name'], choices_shortlist[label_col]))
    # add an NA entry to the dictionary if the parameter was chosen
    if na_include is True:
        mapping_dict['No_data_available_NA'] = 'No data available (NA)'
    # if something was missing from the tool, it will be its own label
    for value in values_list:
        if value not in mapping_dict:
            mapping_dict[value] = value
    # None breaks everything. Trying to change it
    mapping_dict['none'] = 'None '
    # build a list with the labels
    values_list = [mapping_dict.get(value, value) for value in values_list]
    return values_list

def round_sig(x, sig):
    """
    Auxiliary function to round to significant digits
    
    Parameters:
    ----------
        x : float
        sig : integer
        
    Returns:
    -------
        x rounded to sig digits
    """
    if x == 0:
        return 0
    return round(x, sig - int(np.floor(np.log10(abs(x)))) - 1)

def smart_rounding(mean, moe):
    """
    Auxiliary function to do smart round of the mean and median
    
    Parameters:
    ----------
    mean : float
        mean (or median) you'd like to round
    moe : float
        precalculated Margin of Error of the mean (median)
    Returns:
    -------
        list
        rounded mean (median) and rounded MoE
    """
    s = str(abs(moe)).replace('.', '')
    non_zero_digits = [ch for ch in s if ch != '0']
    
    if len(non_zero_digits) > 1:
        first_digit = int(non_zero_digits[0])
        sec_digit = int(non_zero_digits[1])
    else:
        first_digit = int(non_zero_digits[0])
        sec_digit = 0

    if (first_digit == 1) & (sec_digit > 0):
        sig_digits = 2
    else:
        sig_digits = 1
    rounded_moe = round_sig(moe, sig_digits)

    # Determine number of decimal places for final rounding
    moe_order = -int(np.floor(np.log10(rounded_moe)))
    decimals = moe_order if sig_digits == 1 else moe_order+1
    rounded_mean = round(mean, decimals)

    return rounded_mean, rounded_moe

def weighted_mean(df, weight_column, numeric_column, add_moe=True):
    
    """
    Calculates the weighted mean, weighted median, and other summary statistics 
    for a specified numeric column, using weights from a specified column.

    Parameters:
    ----------
    df : pandas.DataFrame
        The DataFrame containing the data.
        
    weight_column : str
        The name of the column in the DataFrame that contains the weights.
        
    numeric_column : str
        The name of the column in the DataFrame that contains the numeric values 
        for which the weighted statistics are calculated.
        
    add_moe : bool, optional
        If True - counts confidence intevral for the mean using t-test and 
        standard deviation obtained from stratification (assuming normal 
        distribution within strata), rounds mean with respect to MoE.
        If False - does not count margin of error, rounds mean to one digit 
        after coma if all inputs are integer and to two digits otherwise. 

    Returns:
    -------
    pandas.Series
        A Series containing the following summary statistics:
        - 'mean': The weighted mean of the `numeric_column`.
        - 'median': The weighted median of the `numeric_column`.
        - 'max': The maximum value of the `numeric_column`.
        - 'min': The minimum value of the `numeric_column`.
        - 'unweighted_count': The number of rows in the DataFrame (unweighted).
        - 'count': The total weight (rounded).
    """
    
    # get the total sum of the numeric column
    weighted_sum = (df[numeric_column] * df[weight_column]).sum()
    # get the total weight value
    total_weight = df[weight_column].sum()
    # get the weighted mean
    weighted_mean_result = weighted_sum / total_weight
    # max and minimum
    weighted_max_result = df[numeric_column].max()
    weighted_min_result = df[numeric_column].min()
    # get the weighted and unweighted counts
    count_w = round(total_weight,0)
    count = df.shape[0]
    
    # sort the values
    sorted_df = df.sort_values(by=numeric_column)
    cum_weights = sorted_df[weight_column].cumsum()
    # get the index of the median numeric value
    median_index = np.searchsorted(cum_weights, total_weight / 2.0)
    # if the median indexed value is squarely in the middle take it
    if cum_weights.iloc[median_index] == total_weight / 2.0 or sorted_df.shape[0] <= 2:
        weighted_median_result = sorted_df.iloc[median_index][numeric_column]
    else:
    # if its not, take the one after the index or the last value (for super small samples)
        min_of_two = np.min([median_index + 1, sorted_df.shape[0]-1])
        weighted_median_result = sorted_df.iloc[min_of_two][numeric_column]
    # calculating MoE of the median
    if add_moe:
        # Estimate density at the median using weighted KDE
        kde = gaussian_kde(df[numeric_column], weights=df[weight_column])
        density_x = np.linspace(min(df[numeric_column],), max(df[numeric_column],), 1000)
        density_y = kde(density_x)
        f_interp = interp1d(density_x, density_y, fill_value="extrapolate")
        fmed = float(f_interp(weighted_median_result))
        # Calculate MoE of the median for 95% CL and rounding median
        moe_median = 1.96 / (np.sqrt(4 * len(df[numeric_column])) * fmed)
        weighted_median_result, moe_median = smart_rounding(weighted_median_result,moe_median)
    else:    
        # rounding median to one digit
        weighted_median_result = round(weighted_median_result, 1)
            
    # calculating MoE of the mean
    if add_moe:
        # group values by strata (weights)
        x = df[numeric_column].to_numpy()
        w = df[weight_column].to_numpy()
        simple_df = pd.DataFrame({'x_value':x,'weight':w})
        strata = simple_df.groupby('weight')
        weighted_vars = []
        weights = []
        # looping throug strata
        for stratum_value, group in strata:
            x_value = group['x_value']
            n = len(x_value)
            if n==1:
                var_x=(np.sum(x_value) - np.sum(w * x) / np.sum(w))**2 # we cannot calculate variance of one value, so it's like this - comparing to overall mean
            else:
                var_x = x_value.var(ddof=1)  # sample variance within stratum
            # Treat the stratum_value as the weight (proxy for population share)
            weight = stratum_value
            weighted_vars.append((weight**2) * var_x / n)  # stratum-level variance contribution
            weights.append(weight)
        # Variance of weighted mean (stratified)
        var_weighted_mean = sum(weighted_vars) / (sum(weights)**2)
        # Calculate MoE of the mean for 95% CL and rounding mean
        moe_mean = 1.96 * np.sqrt(var_weighted_mean)
        weighted_mean_result, moe_mean = smart_rounding(weighted_mean_result, moe_mean)
    else:
        # rounding mean to two digits 
        weighted_mean_result = round(weighted_mean_result, 2)
            
    weighted_max_result = round(weighted_max_result, 1)
    weighted_min_result = round(weighted_min_result, 1)
    if add_moe:
        output = pd.Series({'mean': weighted_mean_result,
                          'moe_mean' : moe_mean,
                          'median':weighted_median_result,
                          'moe_median': moe_median,
                          'max': weighted_max_result,
                          'min': weighted_min_result,
                          'unweighted_count' : count,
                          'count': count_w})
    else:
        output = pd.Series({'mean': weighted_mean_result,
                          'median':weighted_median_result,
                          'max': weighted_max_result,
                          'min': weighted_min_result,
                          'unweighted_count' : count,
                          'count': count_w})
    return output


def get_variable_type(data, variable_name):
    
    """
    Determines the data type of a specified variable/column in a DataFrame and returns it as a string.

    Parameters:
    ----------
    data : pandas.DataFrame
        The DataFrame containing the data.
        
    variable_name : str
        The name of the column for which the data type is being checked.

    Returns:
    -------
    str
        A string representing the type of the variable. The possible return values are:
        - 'string' for object (text) data types.
        - 'integer' for integer data types (int64, int32).
        - 'decimal' for floating-point data types (float64, float32).
    """
    if data[variable_name].dtype == 'object':
        return 'string'
    elif data[variable_name].dtype == 'int64' or data[variable_name].dtype == 'int' or data[variable_name].dtype == 'int32':
        return 'integer'
    elif data[variable_name].dtype == 'float64' or data[variable_name].dtype == 'float' or data[variable_name].dtype == 'float32':
        return 'decimal'


def check_daf_filter(daf, data, filter_daf, tool_survey):
    
    """
    Validates a filter DataFrame against a Data Analysis Framework (DAF) and a dataset, 
    ensuring that the filter conditions are consistent and correctly defined.

    Parameters:
    ----------
    daf : pandas.DataFrame
        The DataFrame containing the Data Analysis Framework (DAF) data.
        
    data : dict of pandas.DataFrame
        A dictionary containing DataFrames for each datasheet in the tool, 
        with datasheet names as keys.
        
    filter_daf : pandas.DataFrame
        The DataFrame containing the filter conditions to be validated against the DAF.
        
    tool_survey : pandas.DataFrame
        The DataFrame containing the survey sheet of your Kobo.

    Raises:
    ------
    ValueError
        If there are any NaN values in the filter DataFrame, if the IDs in 
        the filter DataFrame are not consistent with those in the DAF, 
        or if any filter conditions are incorrectly defined (e.g., variable not found, 
        type mismatches, or disallowed operations for variable types).

    Notes:
    -----
    - The function checks that all filter variables exist in the corresponding 
      datasheet in the data.
    - It also ensures that the variable types match between filter conditions 
      and the corresponding data columns.
    - The allowed operations are validated based on the variable types 
      (e.g., numerical comparisons for numeric types and equality checks for strings).
      
    """
    merged_daf = filter_daf.merge(daf, on='ID', how='inner')
    # some calculate variables can be NaN
    merged_daf = merged_daf.drop(
        ['calculation', 'join', 'disaggregations'], axis=1)
    # check if rows contain NaN
    if filter_daf.isnull().values.any():
        raise ValueError("Some rows in the filter sheet contain NaN")

    # check IDs consistency
    if len(merged_daf) != len(filter_daf):
        raise ValueError("Some IDs in file are not in DAF")

    for row_id, row in merged_daf.iterrows():
        # check that filter variable are in the same sheet in the data
        if row['variable_x'] not in data[row['datasheet']].columns:
            raise ValueError(f"Filter variable {row['variable_x']} not found in {row['datasheet']}")

        value_type = type(row['value'])

        # check whether the value is an another variable
        if row["value"] in tool_survey['name'].tolist():
            # check that the variable is in the same sheet in the data
            if row['value'] not in data[row['datasheet']].columns:
                raise ValueError(f"Filter value {row['value']} not found in {row['datasheet']}")

            # check that the variable and the value have the same type
            if get_variable_type(data[row['datasheet']], row['variable_x']) != get_variable_type(data[row['datasheet']], row['value']):
                raise ValueError(f"Variable {row['variable_x']} and {row['value']} have different types")

            # check that the operation is allowed for the type
            if get_variable_type(data[row['datasheet']], row['value']) == 'string':
                if row['operation'] not in ["!=", "=="]:
                    raise ValueError(
                        f"Operation {row['operation']} not allowed for string variables")
            continue

        if value_type == str:
            # check that the variable and the value have the same type
            if get_variable_type(data[row['datasheet']], row['variable_x']) != 'string':
                raise ValueError(
                    f"Variable {row['variable_x']} has another type then filter value")
            # check that the operation is allowed for the type
            if row["operation"].strip(' ') not in ["!=", "=="]:
                raise ValueError(
                    f"Operation {row['operation']} not allowed for string variables")
        else:
            # check that the variable and the value have the same type
            if get_variable_type(data[row['datasheet']], row['variable_x']) == 'string':
                raise ValueError(
                    f"Variable {row['variable_x']} has another type then filter value")
            # check that the operation is allowed for the type
            if row["operation"].strip(' ') not in ["<", ">", "<=", ">=", "!=", "=="]:
                raise ValueError(
                    f"Operation {row['operation']} not allowed for numeric variables")


def check_daf_consistency(daf, data, sheets, resolve=False):
    
    """
    Checks the consistency of a Data Analysis Framework (DAF) against 
    the corresponding data sheets, verifying that all variables 
    are properly defined and exist in the appropriate data sheets.

    Parameters:
    ----------
    daf : pandas.DataFrame
        The DataFrame representing the Data Analysis Framework, which includes 
        variables and their associated datasheets.

    data : dict of pandas.DataFrame
        A dictionary containing DataFrames for each datasheet in the tool, 
        with datasheet names as keys.

    sheets : list of str
        A list of names of the sheets to be checked for variable existence.

    resolve : bool, optional
        If set to True, the function will print warnings and drop variables 
        from the DAF that are missing or inconsistent instead of raising 
        exceptions. Default is False.

    Returns:
    -------
    pandas.DataFrame
        The updated DAF after performing consistency checks. Variables 
        that were found to be inconsistent may be dropped if `resolve` is True.

    Raises:
    ------
    ValueError
        If there are any missing variables in the DAF or if any variables 
        do not exist in their corresponding data sheets (unless `resolve` is True).

    Warnings:
    ---------
    Warnings will be printed for any intersections found between variables 
    existing in multiple sheets.

    Notes:
    -----
    - This function checks that all variables listed in the DAF have a 
      corresponding datasheet and that they exist in the relevant data sheets.
    - It also checks for any disaggregations defined in the DAF to ensure they 
      are present in the data sheets.
    - The function prints a summary of the number of variables found in each 
      sheet, as well as any issues identified during the consistency checks.
    """

    # check that all variables have a datasheet
    if daf['datasheet'].isnull().values.any():
        if not resolve:
            raise ValueError('the following are missing ' +
                             ','.join(daf[daf['datasheet'].isnull().values]['variable']))
        else:
            print('the following are missing ' +
                  ','.join(daf[daf['datasheet'].isnull().values]['variable']))
            daf.dropna(subset=['datasheet'], inplace=True)

    # check that all variables in daf are in the corresponding data sheets
    for id, row in daf.iterrows():
        if row["variable"] not in data[row["datasheet"]].columns:
            if not resolve:
                raise ValueError(f"Column {row['variable']} not found in {row['datasheet']}")
            else:
                print(f"Column {row['variable']} not found in {row['datasheet']}")
                daf.drop(id, inplace=True)
        # Do some basic cleaning of disaggregations column
        if row["disaggregations"] not in ["overall", ""] and not pd.isna(row['disaggregations']):
            row["disaggregations"] = row["disaggregations"].replace(" ", "")
            disaggregations_list = row["disaggregations"].split(",")
            # check if disaggregations are present in the relevant sheets
            for disaggregations_item in disaggregations_list:
                if disaggregations_item not in data[row["datasheet"]].columns:
                    error_message = f"Disaggregation {disaggregations_item} not found in {row['datasheet']} for variable {row['variable']}"
                    if not resolve:
                        raise ValueError(error_message)
                    else:
                        print(error_message)
                        daf.drop(id, inplace=True)
                        break
        # check if admin columns are present in the appropriate sheets
        if row["admin"] not in data[row["datasheet"]].columns:
            if not resolve:
                raise ValueError(f"admin {row['admin']} not found in {row['datasheet']} for variable {row['variable']}")
            else:
                print(f"admin {row['admin']} not found in {row['datasheet']} for variable {row['variable']}")
                daf.drop(id, inplace=True)

    # check if variables exist in more than one sheet
    sheet_dict = dict()
    for sheet in sheets:
        colnames = data[sheet].columns
        # drop from colnames the ones that are not in daf
        colnames = colnames[colnames.isin(daf['variable'])]
        sheet_dict[sheet] = set(colnames)

    # check and print all intersections
    for sheet1, sheet2 in combinations(sheet_dict.keys(), 2):
        intersection = sheet_dict[sheet1].intersection(sheet_dict[sheet2])
        if len(intersection) > 0:
            if not resolve:
                warnings.warn(f"Intersection between {sheet1} and {sheet2} : {intersection}")
            else:
                print(f"Intersection between {sheet1} and {sheet2} : {intersection}")
                # print("Resolve by removing from DAF the variables that are in both sheets")
                # daf = daf[~daf['variable'].isin(intersection)]

    for sheet in sheets:
        # check that all sheets have variables in daf
        if not sheet_dict[sheet]:
            print(f"WARNING: Sheet {sheet} has no variables in DAF")
        print(f"Sheet {sheet} has {len(sheet_dict[sheet])} variables")

    return daf


def custom_sort_key(value):
        
    """
    A helper function that helps with list sorting

    Parameters:
    ----------
    daf : str
        A string that will be checked

    Returns:
    -------
    str
        If the value is 'Total' makes sure that it'll be the last on the list. Else, it's just a value

    """
    if isinstance(value, str) and 'Total' in value:
        return (1, value)
    else:
        return (0, value)


def make_pivot(table, index_list, column_list, value):
    """
    Creates a pivot table from a given DataFrame.

    Parameters:
    ----------
    table : pandas.DataFrame
        The DataFrame containing the data to be pivoted.

    index_list : list of str
        The column names to use as the new index for the pivot table.

    column_list : list of str
        The column names to use for creating new columns in the pivot table.

    value : str
        The column name whose values will fill the new pivot table.

    Returns:
    -------
    pandas.DataFrame
        A new DataFrame representing the pivot table, with the specified index 
        and columns and filled with the corresponding values.

    Notes:
    -----
    - The resulting pivot table will have a multi-level index if multiple 
      columns are specified in `index_list` or `column_list`.
    - Missing values in the pivot table will be represented as NaN.
    """
    pivot_table = table.pivot_table(index=index_list,
                                    columns=column_list,
                                    values=value).reset_index()
    return pivot_table


def get_color(value):
    """
    Generate a hexadecimal color code based on the input value, transitioning 
    from red (0%) to yellow/orange (50%) to green (100%).

    Parameters:
    ----------
    value : float
        A numeric value between 0 and 1. Values below 0 will return white, 
        while values above 1 will also return white.

    Returns:
    -------
    str
        A hexadecimal string representing the color. The format is 
        `#RRGGBB`, where `RR`, `GG`, and `BB` are the red, green, and blue 
        components of the color, respectively.

    Notes:
    -----
    - Values in the range [0, 0.5] will transition from red to yellow/orange.
    - Values in the range (0.5, 1] will transition from yellow/orange to green.
    - The function returns "FFFFFF" (white) for NaN values or values outside 
      the range of 0 to 1.
    """
    if value <= 1:
        if value <= 0.5:
            # Interpolate between red and yellow/orange
            red = 255
            green = int(510 * value)  # 255 * 2 * value
            blue = 0
        else:
            # Interpolate between yellow/orange and green
            red = int(510 * (1 - value))  # 255 * 2 * (1 - value)
            green = 255
            blue = 0
        return f"{red:02X}{green:02X}{blue:02X}"
    return "FFFFFF"  # White for NaN or values > 1

def col_num_to_excel(col_num):
    
    """
    Convert a zero-based column index to an Excel-style column letter.

    Parameters:
    ----------
    col_num : int
        A zero-based integer representing the column index. For example:
        - 0 corresponds to 'A'
        - 1 corresponds to 'B'
        - 26 corresponds to 'Z'
        - 27 corresponds to 'AA'

    Returns:
    -------
    str
        A string representing the corresponding Excel column letter(s).

    Notes:
    -----
    - The function supports column indices beyond 25, returning multi-letter 
      column names (e.g., 26 returns 'AA', 27 returns 'AB').
    - This function assumes a zero-based index and does not handle negative 
      inputs or non-integer types.
    """
    letters = string.ascii_uppercase
    if col_num < len(letters):
        return letters[col_num]
    else:
        # For columns beyond 'Z', handle multi-letter column names (e.g., 26 -> AA, 27 -> AB)
        return col_num_to_excel(col_num // 26 - 1) + letters[col_num % 26]

# Added here a function to automatically highlight with red rows with sample <30
def construct_result_table(tables_list, file_name, make_pivot_with_strata=False, color_cells=True, color_less30=True, sort_by_total=False, conditional_formating=True):
    
    """
    Constructs an Excel workbook containing pivot tables from a list of DataFrames.

    This function creates a workbook with two sheets: a table of contents that links to the data 
    and a data sheet with pivot tables generated from the provided list of tables.

    Parameters:
    ----------
    tables_list : list of tuples
        A list of tuples where each tuple contains the following elements:
            - DataFrame: The data to be processed and pivoted.
            - str: An identifier for the table.
            - str: A label for the table.
            - str: The result of the table's significance check.

    file_name : str
        The name of the Excel file to be created (should end with '.xlsx').

    make_pivot_with_strata : bool, optional
        If True, creates pivot tables with strata. Default is False.

    color_cells : bool, optional
        If True, applies color formatting to cells based on their values. Default is True.
    color_less30 : bool, optional
        If True, highlights with red rows with sample size <30
    sort_by_total : bool, optional
        If True, sorts the pivot table by total values. Default is False.

    conditional_formating : bool, optional
        If True, applies conditional formatting to the pivot tables. Default is True.

    Returns:
    -------
    None
        The function creates an Excel file and writes the data to it. The output is saved as 
        specified by the `file_name` parameter.

    Notes:
    -----
    - This function uses the `xlsxwriter` library to create the Excel workbook and apply 
      formatting.
    - The function automatically determines the values variable to be used based on the 
      columns present in each DataFrame.
    - It supports various configurations for how data is pivoted and formatted in the 
      resulting Excel file.
    - The workbook will contain two sheets: "Table_of_content" and "Data".
    - Ensure that the input DataFrames are properly structured for the pivoting process to 
      work correctly.

    Example:
    --------
    >>> construct_result_table(tables_list, "output.xlsx", True, True, False, True)
    """
    # create a workbook object
    workbook = xlsxwriter.Workbook(file_name)
    
    # add the sheets
    content_sheet = workbook.add_worksheet("Table_of_content")
    data_sheet = workbook.add_worksheet("Data")
    # Create a counter
    link_idx = 1
    # Define numeric columns
    numeric_output_cols = ['mean','moe_mean','meadian','moe_median','min','max']
    # Define formatting
    percent_format = workbook.add_format({"num_format": "0.00%"})
    round_format = workbook.add_format({"num_format": "0.00"})

    bold = workbook.add_format({'bold': True})

    border_format=workbook.add_format({
                            'border':1,
                            'align':'left',
                            'font_size':10
                           })

    # add columns in the content sheet
    data = ["ID", "Link", "Significance"]
    for col_num, value in enumerate(data):
        content_sheet.write(0, col_num, value)

    for idx, element in enumerate(tables_list):
        table, ID, label, significance = element
        # print(ID)
        cols_tbl = table.columns
        # The colums that will be added in the pivot
        pivot_column_names = {'disaggregations_category_1', 'oblast', 'macroregion'}

        # what is the values_variable? The object that stores the values of the frequency analysis
        if "perc" in cols_tbl:
            values_variable = "perc"
        else:
            # If we have mulptiple frequencies (case of joining) we have a list of values variables
            if any([x.startswith(('perc_','median_','mean_','max_','min_','category_count_')) for x in cols_tbl]):
                values_variable = [x for x in cols_tbl if x.startswith(('perc_','median_','mean_','max_','min_','category_count_'))]
            # category count can exist within the mean tables too and requires different treatment
            # basically mean is a simple table of means
            # count_mean is the count of that table created for count excel files
            elif 'mean' in table.columns and 'category_count' not in table.columns:
               values_variable = "mean"
            elif 'mean' in table.columns and 'category_count'  in table.columns:
                values_variable = 'count_mean'
            else:
                values_variable = 'category_count'
            pivot_columns = []
        # check that columns are present
        pivot_columns = [col for col in pivot_column_names if col in cols_tbl]
        # else:
        #     pivot_columns = []
        
        # if we have multiple disaggregations, only the first one will be considered for pivoting
        columns = [x for x in cols_tbl if ('disaggregations_category_' in x)]
        missed_cols = set(columns).difference(['disaggregations_category_1'])
        if len(missed_cols)>0:
            pivot_columns.extend(list(missed_cols))
            
        # case of perc and counts is the same
        if values_variable in ["perc" ,'category_count']:
            if make_pivot_with_strata:
                # split the frame with and without 'Total' and pivot them.
                if 'Total' in table['admin_category'].values:
                    table_dirty = table[table['admin_category'] == 'Total']
                    table_clean = table[table['admin_category'] != 'Total']
                    
                    pivot_table_dirty = make_pivot(
                        table_dirty, pivot_columns + ["option"], ["admin_category"], values_variable)
                    pivot_table_clean = make_pivot(
                        table_clean, pivot_columns + ["option"], ["admin_category"], values_variable)
                    # merge the pivoted frames
                    pivot_table = pd.merge(pivot_table_clean, pivot_table_dirty[[
                                           'option', 'Total']], on=['option'], how='left')
                else:
                    # else just pivot the frame regularly
                    pivot_table = make_pivot(
                        table, pivot_columns + ["option"], ["admin_category"], values_variable)
            else:
                # The sample count variable is added here
                if 'general_count' in cols_tbl:
                    pivot_columns.append('general_count')
                    
                # get the option values for sorting
                options_column = table["option"].unique()

                # replace general_count NA values for 'Total' rows woth full_count
                table['general_count'] = table['general_count'].fillna(table['full_count'])

                pivot_table = make_pivot(
                    table, pivot_columns + ["admin_category"], ["option"], values_variable)
                
                # sort the values by the custom sort key
                pivot_table = pivot_table.sort_values(
                    by='admin_category', key=lambda x: x.map(custom_sort_key)).reset_index(drop=True)

                mask = pivot_table['admin_category'] == 'Total'
                total_value = pivot_table.loc[~mask, 'general_count'].sum()
                pivot_table.loc[mask, 'general_count'] = total_value
                
                # if sorted by total apply a different sorting algorithm
                if sort_by_total:
                    if "disaggregations_category_1" in pivot_table.columns:
                    # get the total row
                        total_row = pivot_table[
                            (pivot_table['disaggregations_category_1'] == 'Total') & 
                            (pivot_table['admin_category'] == 'Total')
                        ]
                    else:
                        total_row = pivot_table[(pivot_table['admin_category'] == 'Total')]
                    if not total_row.empty:
                        # Sort the results by value
                        total_values = total_row[options_column].iloc[0]

                        column_value_pairs = list(zip(options_column, total_values))
                        sorted_column_value_pairs = sorted(column_value_pairs, key=lambda x: x[1], reverse=True)

                        sorted_columns = [col for col, _ in sorted_column_value_pairs if col in pivot_table.columns]

                        pivot_table_columns = [col for col in pivot_table.columns if col not in sorted_columns]

                        pivot_table = pivot_table[pivot_table_columns + sorted_columns]
                
                # if macroregion is present in the data, sort by that
                if 'macroregion' in pivot_table.columns:
                    pivot_table = pivot_table.sort_values(
                        by='macroregion'
                    )
        # If the values variable is count_mean, we only need the count itself for our tables. 
        elif values_variable == 'count_mean':
            table = table.reset_index(drop = True)
            cols_to_drop = ['ID','variable','admin','disaggregations_1','total_count_perc','min','max','median','moe_median','mean', 'moe_mean', 'full_count']
            cols_to_keep = [i for i in table.columns if i not in cols_to_drop]
            if make_pivot_with_strata:
                pivot_table = make_pivot(table, pivot_columns, ["admin_category"], 'category_count')
            else:
                pivot_table = table[cols_to_keep]
            
            mask = pivot_table['admin_category'] == 'Total'
            total_value = pivot_table.loc[~mask, 'category_count'].sum()
            pivot_table.loc[mask, 'category_count'] = total_value
                
        elif values_variable == 'mean':
            if make_pivot_with_strata:
                # add numeric columns as a single one by melting the frame
                table = table.reset_index()
                ids = pivot_columns+['ID','admin_category']
                value_vars = [col for col in numeric_output_cols if col in cols_tbl]
                table = pd.melt(table, id_vars=ids, value_vars=value_vars)
                # add new columns to pivot
                values_variable = 'value'
                pivot_columns = pivot_columns +['variable']
                pivot_table = make_pivot(table, pivot_columns, ["admin_category"], values_variable)
            else:
                # if it's just a regular table - remove excessive information
                cols_to_drop = ['ID','variable','admin','disaggregations_1','total_count_perc', 'full_count']
                cols_to_keep = [i for i in cols_tbl if i not in cols_to_drop]
                pivot_table = table[cols_to_keep]
        else:
            # check if we're dealing with a count table
            category_count_columns = [x for x in cols_tbl if x.startswith('category_count_')]
            # and keep only count columns
            cols_to_keep = ([x for x in cols_tbl if '_category' in x]
            +(['option']  if 'option' in cols_tbl else [])
            +(category_count_columns if category_count_columns else
            [x for x in cols_tbl if x.startswith(('perc_','median_','mean_','max_','min_'))])
            +[x for x in cols_tbl if x.endswith('_count')])
            # and keep only what we need. No need to pivot if the tables were joined
            pivot_table = table[cols_to_keep]
            
        if 'macroregion' in pivot_table.columns:
            pivot_table = pivot_table.sort_values(by='macroregion')    

        cols_to_drop = ['mean', 'moe_mean','median','moe_median','min','max']
        # drop unnecessary variables if needed
        if values_variable == 'count_mean':
            cols_to_drop = ['mean', 'moe_mean','median','moe_median','min','max']+[x for x in cols_tbl if x.startswith(('median_','mean_','max_','min_'))]
            cols_to_keep = [i for i in pivot_table.columns if i not in cols_to_drop]
            pivot_table = pivot_table[cols_to_keep]
        
        
        # format the tables themselves
        
        cell_id = link_idx
        names_id = cell_id+1
        
        link_idx += len(pivot_table) + 3
        # set the column headers
        column_headers = list(pivot_table.columns)
        
        # write the label
        data_sheet.write(cell_id, 0, label)
        for col_num, header in enumerate(column_headers):
            data_sheet.write(names_id, col_num, header)

        # set the cell formatting to numeric/percentage
        for row_num, row in pivot_table.iterrows():
            for col_num, (column_name, value) in enumerate(row.items()):
                if column_name not in ['disaggregations_category_1', 'disaggregations_category_2', 'admin_category', 'option', 
                            'strata_name', 'raion', 'oblast', 'macroregion',
                            'mean', 'moe_mean', 'median','moe_median', 'max' ,'min', 'disaggregations_category_3'
                            'count','full_count','weighted_count','unweighted_count','category_count','general_count']:
                    if pd.isna(value):
                        data_sheet.write(row_num + 2 +cell_id, col_num, None)
                    elif values_variable not in ["mean", "count_mean", "value", "category_count"]:
                        data_sheet.write(row_num + 2 +cell_id, col_num, value, percent_format)
                    else:
                        data_sheet.write(row_num + 2 +cell_id, col_num, value, round_format)
                
                else:
                    if pd.isna(value):
                        data_sheet.write(row_num + 2 +cell_id, col_num, None)
                    else:
                        data_sheet.write(row_num + 2 +cell_id, col_num, value)
        
        #color code the percentages
        if  values_variable =='perc' or any(str(col).startswith('perc') for col in pivot_table.columns):
            # get the columns to exclude
            exclude_prefixes = ['median_','mean_','max_','min_']
            
            exclude_columns = ['disaggregations_category_1', 'disaggregations_category_2', 'admin_category', 'option', 
                            'strata_name', 'raion', 'oblast', 'macroregion',
                            'mean', 'moe_mean', 'median','moe_median', 'max' ,'min', 'disaggregations_category_3',
                            'count','full_count','weighted_count','unweighted_count','category_count','general_count']
            # get the columns that need to be color coded and formated
            desired_columns = [col for col in pivot_table.columns if col not in exclude_columns or any(col.startswith(prefix) for prefix in exclude_prefixes)]
            # the table column extent
            first_column_index = pivot_table.columns.get_loc(desired_columns[0])
            last_column_index = pivot_table.columns.get_loc(desired_columns[-1])
            # The table row extent
            first_cell = f"{col_num_to_excel(first_column_index)}{names_id+2}"
            last_cell = f"{col_num_to_excel(last_column_index)}{len(pivot_table)+names_id+1}"
            
            # if needed color code and add borders to the table with the 3_color_scale
            if color_cells and conditional_formating:
                data_sheet.conditional_format(f"{first_cell}:{last_cell}", 
                                            {'type':'3_color_scale',
                                            'min_value': 0,
                                            'max_value': 1})
            if conditional_formating:
                data_sheet.conditional_format( f"{first_cell}:{last_cell}" ,
                                    { 'type' : 'no_blanks' ,
                                    'format' : border_format} )
        # Means are formatted differently
        elif  values_variable =='mean' or any(str(col).startswith('mean_') for col in pivot_table.columns):
            # Get the list of relevant columns
            desired_columns =   [col for col in pivot_table.columns if str(col).startswith(('mean_','median_','max_','min_')) or col in ['mean', 'moe_mean','median','moe_median','max','min']]         
            # the table column extent
            first_column_index = pivot_table.columns.get_loc(desired_columns[0])
            last_column_index = pivot_table.columns.get_loc(desired_columns[-1])
            # The table row extent
            first_cell = f"{col_num_to_excel(first_column_index)}{names_id+2}"
            last_cell = f"{col_num_to_excel(last_column_index)}{len(pivot_table)+names_id+1}"

            # If the formatting was required, get the parameters and set it
            if color_cells and conditional_formating:
                for des_col in desired_columns:
                    max_val = max(pivot_table[des_col])
                    min_val = min(pivot_table[des_col])
                    
                    column_index = pivot_table.columns.get_loc(des_col)
                    
                    
                    first_cell_c = f"{col_num_to_excel(column_index)}{names_id+2}"
                    last_cell_c = f"{col_num_to_excel(column_index)}{len(pivot_table)+names_id+1}"
                    data_sheet.conditional_format(f"{first_cell_c}:{last_cell_c}", 
                                    {'type':'3_color_scale',
                                    'min_value': min_val,
                                    'max_value': max_val})

                data_sheet.conditional_format( f"{first_cell}:{last_cell}" ,
                                                { 'type' : 'no_blanks' ,
                                                'format' : border_format} )
                      
        # Format the contents sheet

        if isinstance(values_variable,list):
            link_value = '' #', '.join(values_variable)
        else:
            link_value = values_variable
        # Add the label to the hyperlink
        text_on_link = f"{label} {link_value}"
        if len(text_on_link)>150:
            text_on_link= text_on_link[:150]+'...'
        # add the hyperlink itself
        link_text = f'=HYPERLINK("#\'Data\'!A{cell_id+1}", "{text_on_link}")'
        # write the files
        content_sheet.write(idx+1, 0, ID)
        content_sheet.write(idx+1, 1, link_text,bold)
        content_sheet.write(idx+1, 2, significance)

    # set the width of the first column in the contents sheet
    data_sheet.autofit()
    data_sheet.set_column_pixels(0, 0, 150)
    # set the width of columns in the data sheet
    for col in range(26):
        data_sheet.set_column(col, col, 30)
    
    workbook.close()


def disaggregation_creator(daf_final, data, filter_dictionary, tool_choices, tool_survey,label_colname, check_significance, add_moe, weight_column=None):

    """
    Creates disaggregated data tables based on specified configurations and conditions.

    This function processes a DataFrame and generates frequency counts or 
    numerical summaries based on specified disaggregations. It supports filtering 
    of data based on predefined conditions and can optionally perform significance 
    testing on the results.

    Parameters:
    ----------
    daf_final : pd.DataFrame
        A DAF object containing metadata and configurations for the disaggregation process,
        including column names, calculation types, disaggregation definitions, and other 
        relevant parameters.

    data : dict of pd.DataFrame
        A dictionary where keys are data sheet names and values are DataFrames containing
        the actual survey response data.

    filter_dictionary : dict
        A mapping of question IDs to filter conditions, used to determine which rows of data
        should be included based on specified filters.

    tool_choices : pd.DataFrame
        A tool_choices object created from 'choices' sheet of your Kobo.

    tool_survey : pd.DataFrame
         A tool_survey object created from 'survey' sheet of your Kobo.

    label_colname : str
        The name of the column used for labeling or identifying specific rows in the data.

    check_significance : bool
        A flag indicating whether to perform statistical significance tests on the disaggregated 
        results.
        
    add_moe : bool
        A flag indicating whether to add Margin of Error and smart rounding for mean and median.

    weight_column : str, optional
        The name of the column used for weighting responses in the calculations. If not provided, 
        a default weight of 1 will be assigned.

    Returns:
    -------
    df_list : list
        A list of DataFrames, each containing the disaggregated results based on the specified
        configurations. Each DataFrame corresponds to a DAF row and may include
        frequency counts, means, or other statistical summaries, depending on the input data 
        and calculations specified.


    Notes:
    ------
    - The function expects the DataFrames in the `data` dictionary to have columns as specified
      in `daf_final`. 
    - Ensure that the disaggregation variables are properly defined in `daf_final` to avoid 
      potential errors during processing.
    - The statistical significance tests performed depend on the independent variables defined 
      in the disaggregation settings and will generate p-values to evaluate differences between 
      groups.
    """
    # If no weight_column was specified, the weight will be set to 1 on all sheets
    if weight_column == None:
        for sheet in data:
            data[sheet]['weight'] = 1
        weight_column = 'weight'
    
    # Check if needs to add MoE columns
    if add_moe:
        numeric_output_cols = ['mean','moe_mean','median','moe_median','min','max']
    else:
        numeric_output_cols = ['mean','median','min','max']
    # get only the relevant functions from the daf and split the object in two
    # Separate processes will be applied for numeric and frequency analyses
    daf_final_freq = daf_final[daf_final['func'].isin(
        ['freq', 'select_one', 'select_multiple'])]
    daf_final_num = daf_final[daf_final['func'].isin(['numeric', 'mean'])]

    daf_final_freq.reset_index(inplace=True)
    daf_final_num.reset_index(inplace=True)

    df_list = []

    if len(daf_final_freq) > 0:
        for i, row in daf_final_freq.iterrows():
            # break down the disaggregations into a convenient list
            if not pd.isna(daf_final_freq.iloc[i]['disaggregations']):
                if ',' in daf_final_freq.iloc[i]['disaggregations']:
                    disaggregations = daf_final_freq.iloc[i]['disaggregations'].split(
                        ',')
                else:
                    disaggregations = [
                        daf_final_freq.iloc[i]['disaggregations']]
                disaggregations = [s.replace(" ", "") for s in disaggregations]
            else:
                disaggregations = []
            if not pd.isna(daf_final_freq.iloc[i]['calculation']):
                # break down the calculations (multiple entries are allowed)
                if ',' in daf_final_freq.iloc[i]['calculation']:
                    calc = daf_final_freq.iloc[i]['calculation'].split(',')
                else:
                    calc = [daf_final_freq.iloc[i]['calculation']]
                calc = [x.strip(' ') for x in calc]
            else:
                calc = 'None'

            # If the current ID is in the list of filtered IDs, get the filtering text and subset the dataframe
            if daf_final_freq.iloc[i]['ID'] in filter_dictionary.keys():
                filter_text = 'data["'+daf_final_freq.iloc[i]['datasheet'] + \
                    '"]['+filter_dictionary[daf_final_freq.iloc[i]['ID']]
                data_temp = eval(filter_text)
            else:
                # If not, just keep the regular frame but select the appropriate sheet
                data_temp = data[daf_final_freq.iloc[i]['datasheet']]

            # keep only those columns that we'll need
            selected_columns = [daf_final_freq['variable'][i]] + \
                disaggregations+[daf_final_freq['admin'][i]]+[weight_column]
            # add overall because we will need it in significance analysis
            if 'overall' not in selected_columns:
                selected_columns = selected_columns + ['overall']
            # get the number of rows for calculation of number percentage of respondents who saw this question
            total_nrow = data_temp.shape[0]
            data_temp = data_temp[selected_columns]
            # for some weird cases, replace any + signs with a simple blank
            data_temp.loc[:, daf_final_freq['variable'][i]] = data_temp.loc[:, daf_final_freq['variable'][i]].apply(
                lambda x: re.sub(' +', ' ', x) if isinstance(x, str) else x)
            # if the row specifies inclusion of NAs - replace them with the static text
            if 'include_na' in calc:
                data_temp.loc[:, daf_final_freq['variable'][i]] = data_temp[daf_final_freq['variable'][i]].fillna(
                    'No_data_available_NA')
                # set the variable of NA inclusion to True for our map_names function
                na_includer = True
            else:
                # remove NA rows otherwise
                data_temp = data_temp[data_temp[daf_final_freq['variable'][i]].notna(
                )]
                na_includer = False

            if data_temp.shape[0] > 0:
                # get the total number of respondents who have answered the question
                freq_count = data_temp.shape[0]
            # keep a backup for select multiples
                data_temp_backup = data_temp.copy()

                # break down the data form SM
                if daf_final_freq.iloc[i]['q.type'] in ['select_multiple']:
                    # set up a temp ID column as an alternative for uuid for removal of non-unique select_multiple entries later on
                    data_temp['ID_column'] = data_temp.index
                    # remove excessive spaces
                    data_temp.loc[:, daf_final_freq.iloc[i]['variable']
                                  ] = data_temp[daf_final_freq.iloc[i]['variable']].str.strip()
                    # split into multiple  by ' ' delimiter
                    data_temp.loc[:, daf_final_freq.iloc[i]['variable']
                                  ] = data_temp[daf_final_freq.iloc[i]['variable']].str.split(' ').copy()
                    # Separate rows using explode
                    data_temp = data_temp.explode(
                        daf_final_freq.iloc[i]['variable'], ignore_index=True)
                    # just in case somebody has duplicated entries within the select multiple
                    data_temp.drop_duplicates(inplace = True)
                    data_temp.drop('ID_column', axis =1, inplace = True)
                # which columns are we calculating the frequency for?
                groupby_columns = [daf_final_freq['admin'][i]] + \
                    disaggregations+[daf_final_freq['variable'][i]]
                # check significance if such was specified
                if check_significance ==True:
                    special_mapping = False
                    # check different cases of dependence     
                    # base case         
                    if len(disaggregations)>0:
                        independent_variables = disaggregations
                        admin_variable = daf_final_freq['admin'][i]
                    # if no disaggregations were chosen - test geo dependence if there are multiple groups
                    elif len(disaggregations)==0 and daf_final_freq['admin'][i] not in ['Overall','overall']:
                        independent_variables = [daf_final_freq['admin'][i]]
                        admin_variable = 'overall'
                        special_mapping= True
                    # if not, don't do anything
                    else:
                        independent_variables = None
                        admin_variable = 'overall'
                    admin_ls = data_temp[admin_variable].unique()
                    admin_frame = []
                    # quick variance analysis
                    if independent_variables is not None:
                        admin_ls = [x for x in admin_ls if x is not None]
                        # add a general p-value to perform a base test with no admin disaggregation
                        admin_ls = admin_ls +['general']
                        # objects that'll hold our results
                        p_value_general = 1
                        all_p_values = []
                        # get the variance columns
                        variance_columns = [daf_final_freq['variable'][i]]+independent_variables
                    
                        for adm in admin_ls:
                            # if we aren't dealing with a general case - get the subset of the data
                            if adm != 'general':
                                data_temp_anova = data_temp[data_temp[admin_variable]==adm]
                            else:
                                data_temp_anova = data_temp
                            # get the chi squared of the variables
                            var_frame = data_temp_anova[variance_columns]
                            contingency_table = pd.crosstab(index = var_frame.iloc[:,0].values, columns =[var_frame[col] for col in variance_columns[1:]])
                            if not contingency_table.empty:
                                stat, p_value, dof, expected = chi2_contingency(contingency_table)
                                p_value = round(p_value,3)
                                # save the results. General results get saved no matter what.
                                # admin disaggregated results get saved only if they're significant
                                if adm == 'general':
                                    p_value_general = p_value
                                elif p_value < 0.05 and adm != 'general':
                                    admin_frame = admin_frame + [adm]
                                    all_p_values = all_p_values + [p_value]
                                        
                        admin_frame = [x for x in admin_frame if x is not None]
                        # some bit of code that does nice formatting of the results
                        if len(admin_frame)>0:
                            if ' Overall' in admin_frame:
                                res_frame = f'Significant relationship (pvalue={p_value_general})'
                            else:
                                # get the labels of admins
                                if admin_variable in set(tool_survey['name']):
                                    admin_frame = map_names_ls(admin_variable,admin_frame,tool_survey, tool_choices,label_colname)                        
                                elif special_mapping==True and len(independent_variables)==1:
                                    if independent_variables[0] in set(tool_survey['name']):
                                        admin_frame = map_names_ls(independent_variables[0],admin_frame,tool_survey, tool_choices,label_colname)

                                admin_text = [f'{name} (p_value={value})' for name, value in zip(admin_frame, all_p_values)]
                                res_frame = 'Significant relationship at: '+', '.join(admin_text)
                        else:
                            res_frame = f'Insignificant relationship (pvalue={p_value_general})'
                    else:
                        res_frame = 'Not applicable'
                else:
                    res_frame = ''

                # get the table with basic frequencies
                summary_stats = data_temp.groupby(groupby_columns)[
                    weight_column].agg(['sum', 'count'])

                # get the same stats but for the full subsample (not calculating option samples)
                groupby_columns_ov = [
                    daf_final_freq['admin'][i]]+disaggregations
                
                summary_stats_var_om = data_temp_backup.groupby(
                    groupby_columns_ov)[weight_column].agg(['sum', 'count'])

                summary_stats.reset_index(inplace=True)
                summary_stats_var_om.reset_index(inplace=True)

                # rename them
                summary_stats.rename(
                    columns={'count': 'unweighted_count'}, inplace=True)
                summary_stats.rename(
                    columns={'sum': 'weighted_count'}, inplace=True)
                # How many people in total have answered the question (weighted)
                summary_stats_var_om.rename(
                    columns={'sum': 'general_count'}, inplace=True)
                 # How many people in total have answered the question (unweighted)
                summary_stats_var_om.rename(
                    columns={'count': 'general_count_uw'}, inplace=True)
                # merge the frames together
                summary_stats_full = summary_stats.merge(
                    summary_stats_var_om, on=groupby_columns_ov, how='left')
                
                # rename the category columns to their standard names
                new_column_names = {daf_final_freq['variable'][i]: 'option',
                                    daf_final_freq['admin'][i]: 'admin_category'}

                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        new_column_names[column_name] = f'disaggregations_category_{j+1}'

                summary_stats_full.rename(
                    columns=new_column_names, inplace=True)
                
                # set the categorizing variables
                summary_stats_full['admin'] = daf_final_freq['admin'][i]
                summary_stats_full['variable'] = daf_final_freq['variable'][i]

                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        summary_stats_full[f'disaggregations_{j+1}'] = disaggregations[j]

                # Replace the option names with their labels
                if tool_survey['name'].isin([daf_final_freq.loc[i, 'variable']]).any():
                    summary_stats_full = map_names(column_name='variable',
                                                   column_values_name='option',
                                                   label_col = label_colname,
                                                   summary_table=summary_stats_full,
                                                   tool_survey=tool_survey,
                                                   tool_choices=tool_choices,
                                                   na_include=na_includer)
                #  Replace the disaggregations names with their labels
                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        if disaggregations[j] in set(tool_survey['name']):
                            summary_stats_full = map_names(column_name=f'disaggregations_{j+1}',
                                                           column_values_name=f'disaggregations_category_{j+1}',
                                                           summary_table=summary_stats_full,
                                                           label_col = label_colname,
                                                           tool_survey=tool_survey,
                                                           tool_choices=tool_choices)
                
                # Replace the admin names with their labels
                if tool_survey['name'].isin([daf_final_freq.loc[i, 'admin']]).any():
                    summary_stats_full = map_names(column_name='admin',
                                                   column_values_name='admin_category',
                                                   summary_table=summary_stats_full,
                                                   label_col = label_colname,
                                                   tool_survey=tool_survey,
                                                   tool_choices=tool_choices)

                
                # add proper labels to variable and disaggregations but keep the original names too
                summary_stats_full['variable_orig'] = summary_stats_full['variable']
                summary_stats_full['variable'] = daf_final_freq['variable_label'][i]
                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        disaggregations_labels = daf_final_freq['disaggregations_label'][i]
                        summary_stats_full[f'disaggregations_{j+1}_orig'] = summary_stats_full[f'disaggregations_{j+1}']
                        summary_stats_full[f'disaggregations_{j+1}'] = disaggregations_labels

                # Calculate the percentages of responses
                summary_stats_full['perc'] = round(
                    summary_stats_full['weighted_count']/summary_stats_full['general_count'], 4)
                
                # round them
                summary_stats_full['weighted_count'] = summary_stats_full['weighted_count'].round()
                summary_stats_full['general_count'] = summary_stats_full['general_count'].round()

                # If the user needs the totals (non-disaggregated). They are added here
                if 'add_total' in calc:
                    summary_stats_total = data_temp.groupby(daf_final_freq['variable'][i])[
                        weight_column].agg(['sum','count'])  
                    summary_stats_total.rename(
                        columns={'count': 'unweighted_count'}, inplace=True)
                    summary_stats_total.reset_index(inplace=True)
                    # sometimes weights are wonky. so we're accounting for that
                    summary_stats_total['perc'] = round(
                        summary_stats_total['sum']/data_temp_backup[weight_column].sum(), 4)
                    summary_stats_total['weighted_count'] = summary_stats_total['sum'].copy().round()
                    # add count (n of non-na rows)
                    summary_stats_total['general_count'] = data_temp_backup.shape[0]
                    # drom the sum column
                    summary_stats_total.drop(columns=['sum'], inplace=True)

                    # rename columns
                    new_column_names = {
                        daf_final_freq['variable'][i]: 'option'}
                    summary_stats_total.rename(
                        columns=new_column_names, inplace=True)
                    # add new columns to match the existing format
                    summary_stats_total['admin'] = 'Total'
                    summary_stats_total['admin_category'] = 'Total'
                    summary_stats_total['variable'] = daf_final_freq['variable'][i]

                    # Replace the option names with their labels
                    if tool_survey['name'].isin([daf_final_freq.loc[i, 'variable']]).any():
                        summary_stats_total = map_names(column_name='variable',
                                                        column_values_name='option',
                                                        summary_table=summary_stats_total,
                                                        label_col = label_colname,
                                                        tool_survey=tool_survey,
                                                        tool_choices=tool_choices,
                                                        na_include=na_includer)
                        
                    # add the variable labels and the original names
                    summary_stats_total['variable_orig'] = summary_stats_total['variable']
                    summary_stats_total['variable'] = daf_final_freq['variable_label'][i]
                    
                    # add the disaggregations labels and the original names
                    if disaggregations != []:
                        for j, column_name in enumerate(disaggregations):
                            summary_stats_total[f'disaggregations_{j+1}'] = 'Total'
                            summary_stats_total[f'disaggregations_category_{j+1}'] = 'Total'
                            
                    # merge together with the base tables
                    summary_stats_full = pd.concat(
                        [summary_stats_full, summary_stats_total], ignore_index=True)

                # Add the total count of the dataframe
                summary_stats_full['full_count'] = freq_count
                
                # Create the label object
                if disaggregations != []:
                    label = daf_final_freq.iloc[i]['variable_label']+' broken down by ' + \
                        daf_final_freq.iloc[i]['disaggregations'] + \
                        ' on the admin of '+daf_final_freq.iloc[i]['admin']
                else:
                    label = daf_final_freq.iloc[i]['variable_label'] + \
                        ' on the admin of '+daf_final_freq.iloc[i]['admin']
                # get only the variables we'll be keeping from these tables
                disagg_columns = [
                    col for col in summary_stats_full.columns if col.startswith('disaggregations') and not col.endswith('orig')]
                
                og_columns = [
                    col for col in summary_stats_full.columns if col.endswith('orig')]
                summary_stats_full['ID'] = daf_final_freq.iloc[i]['ID']
                summary_stats_full['total_count_perc'] = round((summary_stats_full['full_count']/total_nrow)*100,2)
                columns = ['ID', 'admin', 'admin_category', 'option',
                            'variable'] + disagg_columns + ['weighted_count','unweighted_count','perc',
                                                            'general_count','general_count_uw', 'full_count','total_count_perc']+ og_columns
                    
                summary_stats_full = summary_stats_full[columns]
                # merge together into a single list
                df_list.append(
                    (summary_stats_full, daf_final_freq['ID'][i], label,res_frame))

    if len(daf_final_num) > 0:
        # Deal with numerics
        for i, row in daf_final_num.iterrows():
            # split the disaggregations into a convenient list
            if not pd.isna(daf_final_num.iloc[i]['disaggregations']):
                if ',' in daf_final_num.iloc[i]['disaggregations']:
                    disaggregations = daf_final_num.iloc[i]['disaggregations'].split(
                        ',')
                else:
                    disaggregations = [
                        daf_final_num.iloc[i]['disaggregations']]
                disaggregations = [s.replace(" ", "") for s in disaggregations]
            else:
                disaggregations = []
                
            if not pd.isna(daf_final_num.iloc[i]['calculation']):
                # account for multiple calculations
                if ',' in daf_final_num.iloc[i]['calculation']:
                    calc = daf_final_num.iloc[i]['calculation'].split(',')
                else:
                    calc = [daf_final_num.iloc[i]['calculation']]
                calc = [x.strip(' ') for x in calc]
            else:
                calc = 'None'

            # If the current ID is in the list of filtered IDs, get the filtering text and subset the dataframe
            if daf_final_num.iloc[i]['ID'] in filter_dictionary.keys():
                filter_text = 'data["'+daf_final_num.iloc[i]['datasheet'] + \
                    '"]['+filter_dictionary[daf_final_num.iloc[i]['ID']]
                data_temp = eval(filter_text)
            else:
                data_temp = data[daf_final_num.iloc[i]['datasheet']]

            # keep only those columns that we'll need
            selected_columns = [daf_final_num['variable'][i]] + \
                disaggregations+[daf_final_num['admin'][i]]+[weight_column]
            # add overall if we need it for the variance analysis
            if 'overall' not in selected_columns:
                selected_columns = selected_columns + ['overall']
            data_temp = data_temp[selected_columns]
            # get the general shape of the frame to know how many entries we could have had
            total_nrow = data_temp.shape[0]
            # drop all NA observations
            data_temp = data_temp[data_temp[daf_final_num['variable'][i]].notna()]

            if data_temp.shape[0] > 0:
                # get the number of answers
                mean_count = data_temp.shape[0]
                # which columns will we be disaggregating by
                groupby_columns = disaggregations+[daf_final_num['admin'][i]]
                # conduct the tests around here
                if check_significance==True:
                    special_mapping = False
                    # check different cases of dependence   
                    # base case         
                    if len(disaggregations)>0:
                        independent_variables = disaggregations
                        admin_variable = daf_final_num['admin'][i]
                    # if no disaggregations were chosen - test geo dependence if there are multiple groups
                    elif len(disaggregations)==0 and daf_final_num['admin'][i] not in ['Overall','overall']:
                        independent_variables = [daf_final_num['admin'][i]]
                        admin_variable = 'overall'
                        special_mapping = True
                     # if not, don't do anything
                    else:
                        independent_variables = None
                        admin_variable = 'overall'
                     # quick variance analysis
                    if independent_variables is not None:
                        variance_columns = [daf_final_num['variable'][i]]+independent_variables
                        admin_ls = data_temp[daf_final_num['admin'][i]].unique()
                        
                        admin_frame = []
                        
                        admin_ls = [x for x in admin_ls if x is not None]
                        # add a general p-value to perform a base test with no admin disaggregation
                        admin_ls = admin_ls +['general']
                        # objects that'll hold our results
                        p_value_general = 1
                        all_p_values = []
                        
                        for adm in admin_ls:
                            # if we aren't dealing with a general case - get the subset of the data
                            if adm != 'general':
                                data_temp_anova = data_temp[data_temp[daf_final_num['admin'][i]]==adm]
                            else:
                                data_temp_anova = data_temp
                            # format model inputs
                            var_list = [daf_final_num['variable'][i]]+independent_variables
                            dep_list = 'C('+')+C('.join(var_list[1:len(var_list)])+')'
                            formula_mod = f'{var_list[0]} ~ {dep_list}'
                            # perform a basic OLS and get the model's p_value
                            model = ols(formula=formula_mod, data = data_temp).fit()
                            p_val = model.f_pvalue
                            p_val = round(p_val,3)
                            # save the results. General results get saved no matter what.
                            # admin disaggregated results get saved only if they're significant
                            if adm =='general':
                                p_value_general=p_val
                            elif p_val<0.05 and adm != 'general':
                                admin_frame = admin_frame + [adm]
                                all_p_values = all_p_values + [p_val]
                                
                        admin_frame = [x for x in admin_frame if x is not None]
                        # some bit of code that does nice formatting of the results
                        if len(admin_frame)>0:
                            
                            if ' Overall' in admin_frame:
                                res_frame_num = f'Significant relationship (pvalue={p_value_general})'
                            else:
                                # get the list of admin labels
                                if admin_variable in set(tool_survey['name']):
                                    admin_frame = map_names_ls(admin_variable,admin_frame,tool_survey, tool_choices,label_colname)
                                elif special_mapping==True and len(independent_variables)==1:
                                    if independent_variables[0] in set(tool_survey['name']):
                                        admin_frame = map_names_ls(independent_variables[0],admin_frame,tool_survey, tool_choices,label_colname)
                                        
                                admin_text = [f'{name} (p_value={value})' for name, value in zip(admin_frame, all_p_values)]
                                res_frame_num = 'Significant relationship at: '+', '.join(admin_text)
                        else:
                            res_frame_num = f'Insignificant relationship (pvalue={p_value_general})'
                    else:
                        res_frame_num = 'Not applicable'
                else:
                    res_frame_num = ''
                
                # get the general disaggregations statistics

                summary_stats = data_temp.groupby(groupby_columns).apply(
                    weighted_mean, weight_column=weight_column, numeric_column=daf_final_num['variable'][i], add_moe=add_moe)

                summary_stats = summary_stats.reset_index()
                
                # standardize the table names
                new_column_names = {
                    daf_final_num['admin'][i]: 'admin_category'}

                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        new_column_names[column_name] = f'disaggregations_category_{j+1}'

                summary_stats.rename(columns=new_column_names, inplace=True)
                # set the original names of the columns
                summary_stats['admin'] = daf_final_num['admin'][i]
                summary_stats['variable'] = daf_final_num['variable'][i]

                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        summary_stats[f'disaggregations_{j+1}'] = disaggregations[j]

                # disaggregations category replacer
                if disaggregations != [] and tool_survey['name'].isin(disaggregations).any():
                    for j, column_name in enumerate(disaggregations):
                        if disaggregations[j] in set(tool_survey['name']):
                            summary_stats = map_names(column_name=f'disaggregations_{j+1}',
                                                      column_values_name=f'disaggregations_category_{j+1}',
                                                      label_col = label_colname,
                                                      summary_table=summary_stats,
                                                      tool_survey=tool_survey,
                                                      tool_choices=tool_choices)

                # admin category replacer
                if tool_survey['name'].isin([daf_final_num.loc[i, 'admin']]).any():
                    summary_stats = map_names(column_name='admin',
                                              column_values_name='admin_category',
                                              label_col = label_colname,
                                              summary_table=summary_stats,
                                              tool_survey=tool_survey,
                                              tool_choices=tool_choices)

                # add the original variable names and labels
                summary_stats['variable_orig'] = summary_stats['variable']
                summary_stats['variable'] = daf_final_num['variable_label'][i]
                if disaggregations != []:
                    for j, column_name in enumerate(disaggregations):
                        disaggregations_labels = daf_final_num['disaggregations_label'][i]
                        summary_stats[f'disaggregations_{j+1}_orig'] = summary_stats[f'disaggregations_{j+1}']
                        summary_stats[f'disaggregations_{j+1}'] = disaggregations_labels
                        
                # if total is added, then it requires a bit different calculations
                if 'add_total' in calc:
                    summary_stats_total = weighted_mean(
                        data_temp, weight_column=weight_column, numeric_column=daf_final_num['variable'][i],add_moe=add_moe).to_frame().transpose()

                    # add new columns to match the existing format
                    summary_stats_total['admin'] = 'Total'
                    summary_stats_total['admin_category'] = 'Total'
                    summary_stats_total['variable'] = daf_final_num['variable'][i]
                    
                    # add original names and values
                    summary_stats_total['variable_orig'] = summary_stats_total['variable']
                    summary_stats_total['variable'] = daf_final_num['variable_label'][i]
                    if disaggregations != []:
                        for j, column_name in enumerate(disaggregations):
                            summary_stats_total[f'disaggregations_{j+1}'] = 'Total'
                            summary_stats_total[f'disaggregations_category_{j+1}'] = 'Total'
                    # merge with the old data
                    summary_stats = pd.concat(
                        [summary_stats, summary_stats_total], ignore_index=True)
                # get the full number of people answering
                summary_stats['full_count'] = mean_count
                # rename the countz
                summary_stats.rename(columns = {'count' : 'weighted_count'}, inplace = True)
                # add the label object
                if disaggregations != []:
                    label = daf_final_num.iloc[i]['variable_label']+' broken down by ' + \
                        daf_final_num.iloc[i]['disaggregations'] + \
                        ' on the admin of '+daf_final_num.iloc[i]['admin']
                else:
                    label = daf_final_num.iloc[i]['variable_label'] + \
                        ' on the admin of '+daf_final_num.iloc[i]['admin']
                summary_stats['total_count_perc'] = round((summary_stats['full_count']/total_nrow)*100,2)
                # get the list of columns we'll be saving
                og_columns = [
                    col for col in summary_stats.columns if col.endswith('orig')]
                disagg_columns = [
                    col for col in summary_stats.columns if col.startswith('disaggregations') and not col.endswith('orig')]
                summary_stats['ID'] = daf_final_num.iloc[i]['ID']
                columns = ['ID', 'admin', 'admin_category', 'variable'] + \
                    disagg_columns + numeric_output_cols + ['weighted_count','unweighted_count' ,'full_count','total_count_perc']+og_columns
                summary_stats = summary_stats[columns]
                # Append the tupple to the list
                df_list.append((summary_stats, daf_final_num['ID'][i], label,res_frame_num))
    return (df_list)

def key_creator(row):
    """
    Creates a unique key string based on the contents of a DataFrame row.

    This function constructs a key used for disaggregating survey data by concatenating various
    attributes from the provided row. The key includes information about question types, 
    options, administrative categories, and disaggregation categories.

    Parameters:
    ----------
    row : pd.Series
        A single row from a DataFrame representing a row of a frequency table, 
        including fields such as question type, options, variable names, 
        administrative information, and disaggregation categories.

    Returns:
    -------
    str
        A formatted key string that uniquely identifies the disaggregation 
        structure for the given row of data. The key is constructed from various 
        components including the question type, variable name, options, 
        administrative information, and disaggregation categories.

    Notes:
    ------
    - The function expects the input `row` to contain specific keys such as 
      'q.type', 'variable_orig', 'admin_orig', and 'admin_category_orig'.
    - The keys for disaggregation categories must follow the naming convention 
      of containing 'disaggregations_category' and ending with 'orig'.
    - Ensure that the input row does not contain missing values in the 
      required fields to avoid generating incomplete keys.
    """
    
    bit_1_gen = 'prop_'+row['q.type'] if row['q.type'] in ['select_one','select_multiple']  else 'mean'
    if 'option_orig' in row.keys():
      bit_2_option = '' if pd.isna(row['option_orig']) else f"%/%{row['option_orig']}"
    else:
        bit_2_option = 'mean'
    bit_3_admin = '@/@'+ row['admin_orig'] + '%/%' + row['admin_category_orig']
  
    cat_dem = [col for col in row.index if 'disaggregations_category' in col and col.endswith('orig')]
    cat_basic = [col for col in row.index if 'category' not in col and col.endswith('orig') and col.startswith('disaggregations')]
    
    combined_disaggs = [f"{row[basic]}%/%{row[dem]}" for basic, dem in zip(cat_basic, cat_dem) if not pd.isna(row[basic]) and not pd.isna(row[dem])]
    bit_4_disaggs = '-/-'.join(combined_disaggs)
    return bit_1_gen +'@/@' +row['variable_orig'] + bit_2_option + bit_3_admin + '-/-'+bit_4_disaggs
    

##################################### wide table construct #####################################

def construct_result_wide_table(tables_list, file_name):
    """
    Processes frequency and numeric tables, pivots them using perc or value values and admin levels, and saves them into an Excel file.
    :param tables_list: List of tuples containing table data, ID, label, and significance
    :param file_name: Output Excel file name
    """
    numeric_list = []
    freq_list = []
    numeric_output_cols = ['mean','moe_mean','meadian','moe_median','min','max']

    for idx, element in enumerate(tables_list):
        table, ID, label, significance = element
        cols_tbl = table.columns
        # The colums that will be added in the pivot
        pivot_column_names = {'disaggregations_category_1', 'oblast', 'macroregion'}

        # what is the values_variable? The object that stores the values of the frequency analysis
        if "perc" in cols_tbl:
            values_variable = "perc"
        else:
            # If we have mulptiple frequencies (case of joining) we have a list of values variables
            if any([x.startswith(('perc_','median_','mean_','max_','min_','category_count_')) for x in cols_tbl]):
                values_variable = [x for x in cols_tbl if x.startswith(('perc_','median_','mean_','max_','min_','category_count_'))]
            # category count can exist within the mean tables too and requires different treatment
            # basically mean is a simple table of means
            # count_mean is the count of that table created for count excel files
            elif 'mean' in table.columns and 'category_count' not in table.columns:
               values_variable = "mean"
            elif 'mean' in table.columns and 'category_count'  in table.columns:
                values_variable = 'count_mean'
            else:
                values_variable = 'category_count'
            pivot_columns = []
        # check that columns are present
        pivot_columns = [col for col in pivot_column_names if col in cols_tbl]
        
        # if we have multiple disaggregations, only the first one will be considered for pivoting
        columns = [x for x in cols_tbl if ('disaggregations_category_' in x)]
        missed_cols = set(columns).difference(['disaggregations_category_1'])
        if len(missed_cols)>0:
            pivot_columns.extend(list(missed_cols))
            
        # case of perc and counts is the same
        if values_variable in ["perc"]:
            table["general_count"] = table["general_count"].fillna(0)
            dissagr_col = [col for col in table.columns if col.startswith("disaggregations_category_")]

            # split the frame with and without 'Total' and pivot them.
            if 'Total' in table['admin_category'].values:
                table_dirty = table[table['admin_category'] == 'Total']
                table_clean = table[table['admin_category'] != 'Total']
                
                pivot_table_dirty = make_pivot(
                    table_dirty, pivot_columns + ["option"], ["admin_category"], values_variable)
                pivot_table_clean = make_pivot(
                    table_clean, pivot_columns + ["option"], ["admin_category"], values_variable)
                # merge the pivoted frames
                pivot_table = pd.merge(pivot_table_clean, pivot_table_dirty[[
                                        'option', 'Total']], on=['option'], how='left')
            else:
                # else just pivot the frame regularly
                pivot_table = make_pivot(
                    table, pivot_columns + ["option"], ["admin_category"], values_variable)
                
        elif values_variable == 'mean':
            mean_table_backup = table.copy()
            # add numeric columns as a single one by melting the frame
            table = table.reset_index()
            ids = pivot_columns+['ID', 'admin_category']
            value_vars = [col for col in numeric_output_cols if col in cols_tbl]
            table = pd.melt(table, id_vars=ids, value_vars=value_vars)
            # add new columns to pivot
            values_variable = 'value'
            pivot_columns = pivot_columns + ['variable']
            pivot_table = make_pivot(table, pivot_columns, ["admin_category"], values_variable)
            
        if 'macroregion' in pivot_table.columns:
            pivot_table = pivot_table.sort_values(by='macroregion')    

        cols_to_drop = ['mean', 'moe_mean','median','moe_median','min','max']
        # drop unnecessary variables if needed
        if values_variable == 'count_mean':
            cols_to_drop = ['mean', 'moe_mean','median','moe_median','min','max']+[x for x in cols_tbl if x.startswith(('median_','mean_','max_','min_'))]
            cols_to_keep = [i for i in pivot_table.columns if i not in cols_to_drop]
            pivot_table = pivot_table[cols_to_keep]

        if values_variable in ["value"]:
            
            dissagr_col = [col for col in mean_table_backup.columns if col.startswith("disaggregations_category_")]
            question_label = mean_table_backup["variable"].values[0]
            if len(dissagr_col) > 0:
                question_label += " by "
            for idx, dis in enumerate(dissagr_col):
                question_label += mean_table_backup[f"disaggregations_{idx + 1}"].values[0] + " "

            pivot_table["question"] = question_label

            numeric_list.append((pivot_table, pivot_table["question"].values[0], mean_table_backup["admin"].values[0]))

        if values_variable in ["perc"]:
            dissagr_col = [col for col in table.columns if col.startswith("disaggregations_category_")]

            question_label = table["variable"].values[0]
            if len(dissagr_col) > 0:
                question_label += " by "
            for idx, dis in enumerate(dissagr_col):
                question_label += table[f"disaggregations_{idx + 1}"].values[0] + " "

            pivot_table["question"] = question_label

            freq_list.append((pivot_table, pivot_table["question"].values[0], table["admin"].values[0]))

        
    def group_and_concat(data_list, prefix):
        grouped = {}
        for pivot_table, variable, admin in data_list:
            if admin not in grouped.keys():
                grouped[admin] = []
            grouped[admin].append(pivot_table)
        
        for admin, tables in grouped.items():
            concated_table = pd.concat(tables, ignore_index=True)
            dissagr_col = [col for col in concated_table.columns if col.startswith("disaggregations_category_")]
            if prefix == "Frequency":
                if "Total" in concated_table.columns:
                    concated_table.rename(columns={"Total": "total_for_question"}, inplace=True)
                    concated_table = concated_table[["question", "option"] + dissagr_col + ["total_for_question"] + [col for col in concated_table.columns if col not in ["question", "option"] + dissagr_col + ["total_for_question"]]]
                else:
                    concated_table = concated_table[["question", "option"] + dissagr_col + [col for col in concated_table.columns if col not in ["question", "option"] + dissagr_col]]
            else:
                if "Total" in concated_table.columns:
                    concated_table.rename(columns={"Total": "total_for_question"}, inplace=True)
                    concated_table = concated_table[["question", "variable"] + dissagr_col + ["total_for_question"] + [col for col in concated_table.columns if col not in ["question", "variable"] + dissagr_col + ["total_for_question"]]]
                else:
                    concated_table = concated_table[["question", "variable"] + dissagr_col + [col for col in concated_table.columns if col not in ["question", "variable"] + dissagr_col]]
            grouped[admin] = concated_table
        
        return grouped

    grouped_freq = group_and_concat(freq_list, prefix="Frequency")
    grouped_numeric = group_and_concat(numeric_list, prefix="Numeric")

    workbook = xlsxwriter.Workbook(file_name)

    def save_to_sheets(grouped_data, workbook, prefix):
        
        for admin, table in grouped_data.items():
            dissagr_col = [col for col in table.columns if col.startswith("disaggregations_category_")]
            table[dissagr_col] = table[dissagr_col].fillna("Total")
            table = table.fillna(0).replace([float('inf'), -float('inf')], "Inf")

            sheet_name = f"{prefix}_{admin}"[:31]
            worksheet = workbook.add_worksheet(sheet_name)

            if len(dissagr_col) > 0:
                last_diss_col_index = table.columns.get_loc(dissagr_col[-1])
                worksheet.set_column(0, last_diss_col_index, 40)
                worksheet.set_column(last_diss_col_index + 1, len(table.columns) - 1, 20)
            else:
                worksheet.set_column(0, 2, 40)
                worksheet.set_column(3, len(table.columns) - 1, 20)

            worksheet.set_row(0, 20)

            header_format = workbook.add_format({
                'bold': True,
                'font_color': 'white',
                'bg_color': 'blue',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            worksheet.write_row(0, 0, table.columns.tolist(), header_format)


            col_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'num_format': '0.00',
            })

            percentage_format = workbook.add_format({
                'num_format': '0.00%',
                'align': 'center',
                'valign': 'vcenter'
            })

            variable_col_index = table.columns.get_loc("question")
            if prefix == "Frequency":
                if len(dissagr_col) > 0:
                    last_dis_col_index = table.columns.get_loc(dissagr_col[-1])
                else:
                    last_dis_col_index = 1
            last_value = None
            merge_start = None

            for row_num, row_data in enumerate(table.values, start=1):
                current_value = row_data[variable_col_index]
                if current_value == last_value:
                    if merge_start is None:
                        merge_start = row_num - 1
                else:
                    if merge_start is not None:
                        worksheet.merge_range(
                            merge_start, 0, row_num - 1, 0,
                            last_value, col_format
                        )
                        merge_start = None
                    last_value = current_value
                
                if prefix == "Frequency":
                    worksheet.write_row(row_num, 0, row_data[:last_dis_col_index + 1], col_format)
                    worksheet.write_row(row_num, last_dis_col_index + 1, row_data[last_dis_col_index + 1:], percentage_format)
                else:
                    worksheet.write_row(row_num, 0, row_data, col_format)
            worksheet.set_row(len(table), 15)

            if merge_start is not None:
                worksheet.merge_range(
                    merge_start, 0, row_num, 0,
                    last_value, col_format
                )

    save_to_sheets(grouped_numeric, workbook, prefix="Numeric")
    save_to_sheets(grouped_freq, workbook, prefix="Frequency")

    workbook.close()


def construct_count_wide_table(tables_list, file_name):
    """
    Processes frequency and numeric tables, pivots them using general_count values and admin levels, and saves them into an Excel file.
    :param tables_list: List of tuples containing table data, ID, label, and significance
    :param file_name: Output Excel file name
    """
    numeric_list = []
    freq_list = []

    for idx, element in enumerate(tables_list):
        table, ID, label, significance = element
        cols_tbl = table.columns
        # The colums that will be added in the pivot
        pivot_column_names = {'disaggregations_category_1', 'oblast', 'macroregion'}

        # what is the values_variable? The object that stores the values of the frequency analysis
        if "perc" in cols_tbl:
            values_variable = "general_count_perc"
        else:
            # If we have mulptiple frequencies (case of joining) we have a list of values variables
            if any([x.startswith(('perc_','median_','mean_','max_','min_','category_count_')) for x in cols_tbl]):
                values_variable = [x for x in cols_tbl if x.startswith(('perc_','median_','mean_','max_','min_','category_count_'))]
            # category count can exist within the mean tables too and requires different treatment
            # basically mean is a simple table of means
            # count_mean is the count of that table created for count excel files
            elif 'mean' in table.columns and 'category_count' not in table.columns:
               values_variable = "general_count_mean"
            elif 'mean' in table.columns and 'category_count'  in table.columns:
                values_variable = 'count_mean'
            else:
                values_variable = 'category_count'
            pivot_columns = []
        # check that columns are present
        pivot_columns = [col for col in pivot_column_names if col in cols_tbl]
        
        # if we have multiple disaggregations, only the first one will be considered for pivoting
        columns = [x for x in cols_tbl if ('disaggregations_category_' in x)]
        missed_cols = set(columns).difference(['disaggregations_category_1'])
        if len(missed_cols)>0:
            pivot_columns.extend(list(missed_cols))
            
        # case of perc and counts is the same
        if values_variable in ["general_count_perc"]:
            table["general_count"] = table["general_count"].fillna(0)
            dissagr_col = [col for col in table.columns if col.startswith("disaggregations_category_")]

            pivot_table = make_pivot(
                    table, pivot_columns, ["admin_category"], "general_count")
            
            if "index" in pivot_table.columns:
                pivot_table = pivot_table.drop(columns=["index"])

            if "Total" in pivot_table.columns:
                pivot_table = pivot_table.drop(columns=["Total"])
                
        elif values_variable == 'general_count_mean':
            mean_table_backup = table.copy()
            # add numeric columns as a single one by melting the frame
            table = table.reset_index()
            ids = pivot_columns+['ID', 'admin_category']
            table = pd.melt(table, id_vars=ids, value_vars=['general_count'])
            # add new columns to pivot
            values_variable = 'value'
            pivot_columns = pivot_columns + ['variable']
            pivot_table = make_pivot(table, pivot_columns, ["admin_category"], values_variable)
            
        if 'macroregion' in pivot_table.columns:
            pivot_table = pivot_table.sort_values(by='macroregion')

        if values_variable in ["value"]:
            
            dissagr_col = [col for col in mean_table_backup.columns if col.startswith("disaggregations_category_")]
            question_label = mean_table_backup["variable"].values[0]
            if len(dissagr_col) > 0:
                question_label += " by "
            for idx, dis in enumerate(dissagr_col):
                question_label += mean_table_backup[f"disaggregations_{idx + 1}"].values[0] + " "

            pivot_table["question"] = question_label
            if "variable" in pivot_table.columns:
                pivot_table = pivot_table.drop(columns=["variable"])

            if "Total" in pivot_table.columns:
                pivot_table = pivot_table.drop(columns=["Total"])

            numeric_list.append((pivot_table, pivot_table["question"].values[0], mean_table_backup["admin"].values[0]))

        if values_variable in ["general_count_perc"]:
            dissagr_col = [col for col in table.columns if col.startswith("disaggregations_category_")]

            question_label = table["variable"].values[0]
            if len(dissagr_col) > 0:
                question_label += " by "
            for idx, dis in enumerate(dissagr_col):
                question_label += table[f"disaggregations_{idx + 1}"].values[0] + " "

            pivot_table["question"] = question_label

            freq_list.append((pivot_table, pivot_table["question"].values[0], table["admin"].values[0]))

        
    def group_and_concat(data_list, prefix):
        grouped = {}
        for pivot_table, variable, admin in data_list:
            if admin not in grouped.keys():
                grouped[admin] = []
            grouped[admin].append(pivot_table)
        
        for admin, tables in grouped.items():
            concated_table = pd.concat(tables, ignore_index=True)
            dissagr_col = [col for col in concated_table.columns if col.startswith("disaggregations_category_")]
            if prefix == "Frequency":
                concated_table = concated_table[["question"] + dissagr_col + [col for col in concated_table.columns if col not in ["question"] + dissagr_col]]
            else:
                concated_table = concated_table[["question"] + dissagr_col + [col for col in concated_table.columns if col not in ["question"] + dissagr_col]]
            grouped[admin] = concated_table
        
        return grouped

    grouped_freq = group_and_concat(freq_list, prefix="Frequency")
    grouped_numeric = group_and_concat(numeric_list, prefix="Numeric")

    workbook = xlsxwriter.Workbook(file_name)

    def save_to_sheets(grouped_data, workbook, prefix):
        
        for admin, table in grouped_data.items():
            dissagr_col = [col for col in table.columns if col.startswith("disaggregations_category_")]
            table[dissagr_col] = table[dissagr_col].fillna("Total")
            table = table.fillna(0).replace([float('inf'), -float('inf')], "Inf")

            sheet_name = f"{prefix}_{admin}"[:31]
            worksheet = workbook.add_worksheet(sheet_name)

            if len(dissagr_col) > 0:
                last_diss_col_index = table.columns.get_loc(dissagr_col[-1])
                worksheet.set_column(0, last_diss_col_index, 40)
                worksheet.set_column(last_diss_col_index + 1, len(table.columns) - 1, 20)
            else:
                worksheet.set_column(0, 1, 40)
                worksheet.set_column(1, len(table.columns) - 1, 20)

            worksheet.set_row(0, 20)

            header_format = workbook.add_format({
                'bold': True,
                'font_color': 'white',
                'bg_color': 'blue',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            worksheet.write_row(0, 0, table.columns.tolist(), header_format)


            col_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'num_format': '0',
            })

            variable_col_index = table.columns.get_loc("question")
            if prefix == "Frequency":
                if len(dissagr_col) > 0:
                    last_dis_col_index = table.columns.get_loc(dissagr_col[-1])
                else:
                    last_dis_col_index = 1
            last_value = None
            merge_start = None

            for row_num, row_data in enumerate(table.values, start=1):
                current_value = row_data[variable_col_index]
                if current_value == last_value:
                    if merge_start is None:
                        merge_start = row_num - 1
                else:
                    if merge_start is not None:
                        worksheet.merge_range(
                            merge_start, 0, row_num - 1, 0,
                            last_value, col_format
                        )
                        merge_start = None
                    last_value = current_value
                
                if prefix == "Frequency":
                    worksheet.write_row(row_num, 0, row_data[:last_dis_col_index + 1], col_format)
                    worksheet.write_row(row_num, last_dis_col_index + 1, row_data[last_dis_col_index + 1:], col_format)
                else:
                    worksheet.write_row(row_num, 0, row_data, col_format)
            worksheet.set_row(len(table), 15)

            if merge_start is not None:
                worksheet.merge_range(
                    merge_start, 0, row_num, 0,
                    last_value, col_format
                )

    save_to_sheets(grouped_numeric, workbook, prefix="Numeric")
    save_to_sheets(grouped_freq, workbook, prefix="Frequency")

    workbook.close()


def construct_wide_count_table(disaggregations_perc_new, filename):
    """
    Wrapper for wide tables calls
    """
    construct_result_wide_table(disaggregations_perc_new, filename)
    construct_count_wide_table(disaggregations_perc_new, filename.split('.')[0] + "_count" + ".xlsx")
