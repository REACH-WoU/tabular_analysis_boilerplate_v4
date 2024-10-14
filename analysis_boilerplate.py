import pandas as pd
from pandas.api.types import is_numeric_dtype
import os
from datetime import datetime
from copy import deepcopy


# set up your working directory
os.chdir('/Users/reach/Desktop/Git/tabular_analysis_boilerplate_v4/')

# Read the functions
from src.functions import *

# this is where you input stuff #

# set the parameters and paths
research_cycle = 'test_cycle' # the name of your research cycle
id_round = '1' # the round of your research cycle
date = datetime.today().strftime('%Y_%m_%d')

parquet_inputs = True # Whether you've transformed your data into a parquet inputs
excel_path_data = 'data/test_frame.xlsx' # path to your excel datafile (you may leave it blank if working with parquet inputs)
parquet_path_data = 'data/parquet_inputs/' # path to your parquet datafiles (you may leave it blank if working with excel input)

excel_path_daf = 'resources/inclusion_ccia_overlaps.xlsx' # the path to your DAF file
excel_path_tool = 'resources/MSNA_2024_Kobo_tool_F2F.xlsx' # the path to your kobo tool

label_colname = 'label::English' # the name of your label::English column. Must be identical in Kobo tool and survey sheets!
weighting_column = 'weight' # add the name of your weight column or write None (no quotation marks around None, pls) if you don't have one

sort_by_total = False # Sort choices columns for categorical by "Total" values

sign_check = True # should the script check the significance of the tables?
color_add = True  # should the final output have colored cells?
# end of the input section #

# load the frames
if parquet_inputs:
  files = os.listdir(parquet_path_data)
  files = [file for file in files if file.endswith('.parquet')] # keep only parquet files
  sheet_names = [filename.split('.')[0] for filename in files] # get sheet names
  if len(files)==0:
    raise ValueError('No files in the provided directory')
  data = {}
  for inx,file_id in enumerate(files):
    data[sheet_names[inx]] = pd.read_parquet(os.path.join(parquet_path_data, file_id), engine='pyarrow') # read them
else:
  data = pd.read_excel(excel_path_data, sheet_name=None)

sheets = list(data.keys())

if 'main' not in sheets:
  raise ValueError('One of your sheets (primary sheet) has to be called `main`, please fix.')

tool_choices = load_tool_choices(filename_tool = excel_path_tool,label_colname=label_colname)
tool_survey = load_tool_survey(filename_tool = excel_path_tool,label_colname=label_colname)


# data transformation section below

# add the Overall column to your data
for sheet_name in sheets:
  data[sheet_name]['overall'] =' Overall'
  data[sheet_name]['Overall'] =' Overall'


# check DAF for potential issues
print('Checking Daf for issues')
daf = pd.read_excel(excel_path_daf, sheet_name="main")
daf['variable_label'] = daf['variable_label'].str.replace('"', "'")
daf['disaggregations_label'] = daf['disaggregations_label'].str.replace('"', "'")

colnames_daf = set(['ID','variable','variable_label',
                    'calculation','func','admin','disaggregations','disaggregations_label','join'])

if not colnames_daf.issubset(daf.columns):
  raise ValueError(f'Missing one or more columns from the DAF file main sheet:'+
                                  ', '.join(colnames_daf.difference(daf.columns)))


# remove spaces
for column in ['variable','admin','calculation','func','disaggregations']:
  daf[column] = daf[column].apply(lambda x: x.strip() if isinstance(x, str) else x)


if any(daf['variable']==daf['disaggregations']):
  problematic_ids_str = ', '.join(str(id) for id in daf.loc[daf['variable'] == daf['disaggregations'], 'ID'])
  raise ValueError(f'Variable and disaggregation are duplicated, problematic IDs: ' + \
    problematic_ids_str)



wrong_functions = set(daf['func'])-{'mean','numeric','select_one','select_multiple','freq'}
if len(wrong_functions)>0:
  raise ValueError(f'Wrong functions entered: '+str(wrong_functions)+'. Please fix your function entries')

# add the datasheet column

names_data= pd.DataFrame()

for sheet_name in sheets:
  # get all the names in your dataframe list
  variable_names = data[sheet_name].columns
  # create a lil dataframe of all variables in all sheets
  dat = {'variable' : variable_names, 'datasheet' :sheet_name}
  dat = pd.DataFrame(dat)
  names_data = pd.concat([names_data, dat], ignore_index=True)


names_data = names_data.reset_index(drop=True)
# check if we have any duplicates
duplicates_frame = names_data.duplicated(subset='variable', keep=False)
if duplicates_frame[duplicates_frame==True].shape[0] >0:
  # get non duplicate entries
  names_data_non_dupl = names_data[~duplicates_frame]
  deduplicated_frame = pd.DataFrame()
  # run a loop for all duplicated names
  for i in names_data.loc[duplicates_frame,'variable'].unique():
    temp_names =  names_data[names_data['variable']==i]
    temp_names = temp_names.reset_index(drop=True)
    # if the variable is present in main sheet, keep only that version
    if temp_names['datasheet'].isin(['main']).any():
      temp_names = temp_names[temp_names['datasheet']=='main']
    # else, keep whatever is available on the first row
    else:
      temp_names = temp_names[:1]
    deduplicated_frame=pd.concat([deduplicated_frame, temp_names])
  names_data = pd.concat([names_data_non_dupl,deduplicated_frame])

daf_merged = daf.merge(names_data,on='variable', how = 'left')

daf_merged = check_daf_consistency(daf_merged, data, sheets, resolve=False)

IDs = daf_merged['ID'].duplicated()
if any(IDs):
  raise ValueError('Duplicate IDs in the ID column of the DAF')

# check if DAF numerics are really numeric
daf_numeric = daf_merged[daf_merged['func'].isin(['numeric', 'mean'])]
if daf_numeric.shape[0]>0:
  for i, daf_row in daf_numeric.iterrows():
    res  = is_numeric_dtype(data[daf_row['datasheet']][daf_row['variable']])
    if res == False:
      raise ValueError(f"Variable {daf_row['variable']} from datasheet {daf_row['datasheet']} is not numeric, but you want to apply a mean function to it in your DAF")


print('Checking your filter page and building the filter dictionary')

filter_daf = pd.read_excel(excel_path_daf, sheet_name="filter")

if filter_daf.shape[0]>0:
  check_daf_filter(daf =daf_merged, data = data,filter_daf=filter_daf, tool_survey=tool_survey, tool_choices=tool_choices)
  # Create filter dictionary object 
  filter_daf_full = filter_daf.merge(daf_merged[['ID','datasheet']], on = 'ID',how = 'left')

  filter_dict = {}
  # Iterate over DataFrame rows
  for index, row in filter_daf_full.iterrows():
    # If the value is another variable, don't use the string bit for it
    if isinstance(row['value'], str) and row['value'] in data[row['datasheet']].columns:
      condition_str = f"(data['{row['datasheet']}']['{row['variable']}'] {row['operation']} data['{row['datasheet']}']['{row['value']}'])"
    # If the value is a string and is equal
    elif isinstance(row['value'], str) and row['operation']=='==':
      condition_str = f"(data['{row['datasheet']}']['{row['variable']}'].astype(str).str.contains('{row['value']}', regex=True))"
    # If the value is a string and is not equal
    elif isinstance(row['value'], str) and row['operation']=='!=':
      condition_str = f"(~data['{row['datasheet']}']['{row['variable']}'].astype(str).str.contains('{row['value']}', regex=True))"
    # Otherwise just keep as is
    else:
      condition_str = f"(data['{row['datasheet']}']['{row['variable']}'] {row['operation']} {row['value']})"
    if row['ID'] in filter_dict:
      filter_dict[row['ID']].append(condition_str)
    else:
      filter_dict[row['ID']] = [condition_str]

  # Join the similar conditions with '&'
  for key, value in filter_dict.items():
    filter_dict[key] = ' & '.join(value)
  filter_dict = {key: f'{value}]' for key, value in filter_dict.items()}
else:
  filter_dict = {}


# Check the weights just in case
if weighting_column in ['None','none']:
  weighting_column = None

# Chekc if there's an issue with any of the weights
if weighting_column is not None:
  for sheet_name in sheets:
    if data[sheet_name][weighting_column].isnull().sum().any():
      raise ValueError(f"The weight column in sheet {sheet_name} contains NAs please fix this")
  


# Get the disagg tables

print('Building basic tables')
daf_final = daf_merged.merge(tool_survey[['name','q.type']], left_on = 'variable',right_on = 'name', how='left')
daf_final['q.type']=daf_final['q.type'].fillna('select_one')
disaggregations_full = disaggregation_creator(daf_final, data,filter_dict, tool_choices, tool_survey, label_colname = label_colname, check_significance= sign_check, weight_column =weighting_column)


disaggregations_orig = deepcopy(disaggregations_full) # analysis key table

for element in disaggregations_full:
  if isinstance(element[0], pd.DataFrame):  
    if all(column in element[0].columns for column in element[0].columns if column.endswith('orig')):
      element[0].drop(columns=[col for col in  element[0].columns if col.endswith('orig')], inplace=True)

disaggregations_perc = deepcopy(disaggregations_full) # percentage table
disaggregations_count = deepcopy(disaggregations_full) # count table
disaggregations_count_w = deepcopy(disaggregations_full) # weighted count table

# remove counts prom perc table
for element in disaggregations_perc:
  if isinstance(element[0], pd.DataFrame):  
    columns_to_drop = ['category_count', 'weighted_count', 'unweighted_count']
    # Drop each column if it exists in the DataFrame
    for column in columns_to_drop:
      if column in element[0].columns:
        element[0].drop(columns=column, inplace=True)

# remove perc columns from weighted count table
for element in disaggregations_count_w:
  if isinstance(element[0], pd.DataFrame):  
    columns_to_drop = ['perc', 'unweighted_count']
    for column in columns_to_drop:
      if column in element[0].columns:
        element[0].drop(columns=column, inplace=True)
    element[0].rename(columns={'weighted_count': 'category_count'}, inplace=True)
          
# remove perc columns from unweighted count table
for element in disaggregations_count:
  if isinstance(element[0], pd.DataFrame):  
    columns_to_drop = ['perc', 'weighted_count']
    for column in columns_to_drop:
      if column in element[0].columns:
        element[0].drop(columns=column, inplace=True)
    element[0].rename(columns={'unweighted_count': 'category_count'}, inplace=True)


# Get the columns for Analysis key table 
concatenated_df_orig = pd.concat([tpl[0] for tpl in disaggregations_orig], ignore_index = True)
if 'disaggregations_category_1' in concatenated_df_orig.columns:
  concatenated_df_orig = concatenated_df_orig[(concatenated_df_orig['admin'] != 'Total') & (concatenated_df_orig['disaggregations_category_1'] != 'Total')]
else:
  concatenated_df_orig = concatenated_df_orig[(concatenated_df_orig['admin'] != 'Total')]
  
disagg_columns_og = [col for col in concatenated_df_orig.columns if col.startswith('disaggregations') and not col.endswith('orig')]
ls_orig = ['admin','admin_category','option', 'variable']+disagg_columns_og

for column in ls_orig:
  if column in concatenated_df_orig.columns:
    if column+'_orig' not in concatenated_df_orig.columns:
      concatenated_df_orig[column+'_orig'] = concatenated_df_orig[column]
    concatenated_df_orig[column+'_orig'] = concatenated_df_orig[column+'_orig'].infer_objects(copy=False).fillna(concatenated_df_orig[column])


concatenated_df_orig = concatenated_df_orig.merge(daf_final[['ID','q.type']], on='ID', how='left')

concatenated_df_orig['key'] = concatenated_df_orig.apply(key_creator, axis=1)

if 'mean' in concatenated_df_orig.columns:
  if 'perc' in concatenated_df_orig.columns:
    concatenated_df_orig['perc'] = concatenated_df_orig['perc'].fillna(concatenated_df_orig['mean'])
  else:
    concatenated_df_orig['perc'] = concatenated_df_orig['mean']

concatenated_df_orig=concatenated_df_orig[['key','perc']]

# prepare dashboard inputs 
concatenated_df = pd.concat([tpl[0] for tpl in disaggregations_perc], ignore_index = True)
if 'disaggregations_category_1' in concatenated_df.columns:
  concatenated_df = concatenated_df[(concatenated_df['admin'] != 'Total') & (concatenated_df['disaggregations_category_1'] != 'Total')]
else:
    concatenated_df = concatenated_df[(concatenated_df['admin'] != 'Total')]


disagg_columns = [col for col in concatenated_df.columns if col.startswith('disaggregations')]
concatenated_df.loc[:,disagg_columns] = concatenated_df[disagg_columns].fillna(' Overall')

# Join tables if needed
print('Joining tables if such was specified')
disaggregations_perc_new = deepcopy(disaggregations_perc)
disaggregations_count_new = deepcopy(disaggregations_count)
disaggregations_count_w_new = deepcopy(disaggregations_count_w)

for data_frame in [disaggregations_perc_new,disaggregations_count_new,disaggregations_count_w_new]:
# check if any joining is needed
  if pd.notna(daf_final['join']).any():

    # get other children here
    child_rows = daf_final[pd.notna(daf_final['join'])]

    if any(child_rows['ID'].isin(child_rows['join'])):
      raise ValueError('Some of the join tables are related to eachother outside of their relationship with the parent row. Please fix this')
    

    for index, child_row in child_rows.iterrows():
      child_index = child_row['ID']
      
      if child_index not in daf_final['ID'].values:
        raise ValueError(f'The specified parent index in join column for child row ID = {child_index} doesnt exist in the DAF file')
      
      parent_row = daf_final[daf_final['ID'].isin(child_row[['join']])]
      parent_index = parent_row.iloc[0]['ID']


      # check that the rows are idential
      parent_check = parent_row[['disaggregations','func','calculation','admin','q.type']].reset_index(drop=True)
      child_check = child_row.to_frame().transpose()[['disaggregations','func','calculation','admin','q.type']].reset_index(drop=True)

      # transform None to be of the same type
      parent_check = parent_check.infer_objects(copy=False).fillna('I am empty')
      child_check = child_check.infer_objects(copy=False).fillna('I am empty')
      
      check_result = child_check.equals(parent_check)
      if not check_result:
        raise ValueError(f"Joined rows (parent: {str(parent_row['ID'].values)} and child: {str(child_row['ID'])}) are not identical in terms of admin, calculations, function and disaggregations")
      # get the data and dataframe indeces of parents and children
      child_tupple = [(i,tup) for i, tup in enumerate(data_frame) if tup[1] == child_index]
      parent_tupple = [(i, tup) for i, tup in enumerate(data_frame) if tup[1] == parent_index]

      child_tupple_data = child_tupple[0][1][0].copy()
      child_tupple_index = child_tupple[0][0]
      parent_tupple_data = parent_tupple[0][1][0].copy()
      parent_tupple_index = parent_tupple[0][0]
      # rename the data so that they are readable
      
      
      if parent_tupple_data['variable'][0] == child_tupple_data['variable'][0]:
        var_parent = parent_tupple_data['variable'][0] + '_' +str(parent_tupple_data['ID'][0])
        var_child = child_tupple_data['variable'][0] + '_' + str(child_tupple_data['ID'][0])
        warnings.warn("Some of the rows you're joining have the same variable label. This won't look nice")
      else:
        var_parent = parent_tupple_data['variable'][0] 
        var_child = child_tupple_data['variable'][0] 
      
      varnames = [var_parent,var_child]
      dataframes =[parent_tupple_data, child_tupple_data]

      for var, dataframe in  zip(varnames, dataframes):
        rename_dict = {'mean': 'mean_'+var,'median': 'median_'+var ,'count': 'count_'+var, 
                       'weighted_count': 'weighted_count_'+var,'unweighted_count': 'unweighted_count_'+var,
                       'category_count': 'category_count_'+var,
                      'perc': 'perc_'+var,'min': 'min_'+var, 'max': 'max_'+var}

        for old_name, new_name in rename_dict.items():
          if old_name in dataframe.columns:
            dataframe.rename(columns={old_name: new_name},inplace=True)


      # get the lists of columns to keep and merge
      columns_to_merge = [item for item in parent_tupple_data.columns if 'disaggregations' in item  or 'admin' in item]
      if 'option' in  parent_tupple_data.columns:
        columns_to_merge=columns_to_merge+['option']
        
      columns_to_keep = columns_to_merge+ list(rename_dict.values())

      parent_tupple_data= parent_tupple_data.merge(
        child_tupple_data[child_tupple_data.columns.intersection(columns_to_keep)], 
        on = columns_to_merge,how='left')


      parent_index_f = parent_tupple[0][1][1]

      parent_label_f = str(parent_tupple[0][1][2])
        
      if str(child_tupple[0][1][3]) != '':
        parent_sig_f = str(child_tupple[0][1][3])+' & '+ str(parent_tupple[0][1][3])
      else:
        parent_sig_f = ''

      new_list = (parent_tupple_data,parent_index_f,parent_label_f,parent_sig_f)

      data_frame[parent_tupple_index] = new_list
      del data_frame[child_tupple_index]

# write excel files
print('Writing files')
filename = research_cycle+'_'+id_round+'_'+date

filename_dash = 'output/'+filename+'_dashboard.xlsx'
filename_key = 'output/'+filename+'_analysis_key.xlsx'
filename_toc = 'output/'+filename+'_TOC.xlsx'
filename_toc_count = 'output/'+filename+'_TOC_count_unweighted.xlsx'
filename_toc_count_w = 'output/'+filename+'_TOC_count_weighted.xlsx'
filename_wide_toc = 'output/'+filename+'_wide_TOC.xlsx'


# reorder the tupples
disaggregations_perc_new = sorted(disaggregations_perc_new, key=lambda x: x[1])
disaggregations_count_w_new = sorted(disaggregations_count_w_new, key=lambda x: x[1])
disaggregations_count_new = sorted(disaggregations_count_new, key=lambda x: x[1])

# construct the tables now
construct_result_table(disaggregations_perc_new, filename_toc,make_pivot_with_strata = False, color_cells= color_add, sort_by_total=sort_by_total)
if weighting_column != None:
  construct_result_table(disaggregations_count_w_new, filename_toc_count_w,make_pivot_with_strata = False)
construct_result_table(disaggregations_count_new, filename_toc_count,make_pivot_with_strata = False)
construct_result_table(disaggregations_perc_new, filename_wide_toc,make_pivot_with_strata = True)
concatenated_df.to_excel(filename_dash, index=False)
concatenated_df_orig.to_excel(filename_key, index=False)
print('All done. Congratulations')
