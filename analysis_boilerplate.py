import pandas as pd
import os
from datetime import datetime

# set up your working directory
os.chdir('/Users/reach/Desktop/Git/tabular_analysis_boilerplate_v4/')

# Read the functions
from src.functions import *

# this is where you input stuff #

# set the parameters and paths
research_cycle = 'test_cycle'
id_round = '1'
date = datetime.today().strftime('%Y_%m_%d')

parquet_inputs = True
excel_path_data = 'data/test_frame_2.xlsx'
parquet_path_data = 'data/parquet_inputs/'

excel_path_daf = 'resources/UKR_MSNA_MSNI_DAF_inters.xlsx'
excel_path_tool = 'resources/MSNA_2023_Questionnaire_Final_CATI_cleaned.xlsx'

label_colname = 'label::English'
weighting_column = 'weight' # add the name of your weight column or write None (no quotation marks around None, pls) if you don't have one

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

tool_choices = load_tool_choices(filename_tool = excel_path_tool,label_colname=label_colname)
tool_survey = load_tool_survey(filename_tool = excel_path_tool,label_colname=label_colname)


# data transformation section below

# add the Overall column to your data
for sheet_name in sheets:
  data[sheet_name]['Overall'] ='Overall'


# check DAF for potential issues
print('Checking Daf for issues')
daf = pd.read_excel(excel_path_daf, sheet_name="main")
daf.rename({'admin':'admin'},inplace=True)
filter_daf = pd.read_excel(excel_path_daf, sheet_name="filter")
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
#names_data

daf_merged = daf.merge(names_data,on='variable', how = 'left')

daf_merged = check_daf_consistency(daf_merged, data, sheets, resolve=False)

print('Checking your filter page and building the filter dictionary')

if filter_daf.shape[0]>0:
  check_daf_filter(daf =daf_merged, filter_daf=filter_daf, tool_survey=tool_survey, tool_choices=tool_choices)

  # Create filter dictionary object 
  filter_daf_full = filter_daf.merge(daf_merged[['ID','datasheet']], on = 'ID',how = 'left')

  filter_dict = {}
  # Iterate over DataFrame rows
  for index, row in filter_daf_full.iterrows():
    if isinstance(row['value'], str) and row['value'] in data[row['datasheet']].columns:
      # If the value is another variable, don't use the string bit for it
      condition_str = f"(data['{row['datasheet']}']['{row['variable']}'] {row['operation']} data['{row['datasheet']}']['{row['value']}'])"
    elif isinstance(row['value'], str):
      # If the value is a string add quotes
      condition_str = f"(data['{row['datasheet']}']['{row['variable']}'].astype(str).str.contains('{row['value']}', regex=True))"
    else:
      # Otherwise just keep as is
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

# Get the disagg tables

print('Building basic tables')
daf_final = daf_merged.merge(tool_survey[['name','q.type']], left_on = 'variable',right_on = 'name', how='left')
daf_final['q.type']=daf_final['q.type'].fillna('select_one')
test = disaggregation_creator(daf_final, data,filter_dict, tool_choices, tool_survey, weight_column =weighting_column)

###Get the dashboard inputs

concatenated_df = pd.concat([tpl[0] for tpl in test], ignore_index = True)
concatenated_df = concatenated_df[(concatenated_df['admin'] != 'Total') & (concatenated_df['disaggregations_category_1'] != 'Total')]

concatenated_df.loc[:,['disaggregations_category_1','disaggregations_1']].fillna('Overall', inplace=True)
concatenated_df.loc[:,['disaggregations_category_1', 'disaggregations_1']] = concatenated_df.loc[:,['disaggregations_category_1', 'disaggregations_1']].fillna('Overall')


# Join tables if needed
print('Joining tables if such was specified')
test_new = test.copy()
# check if any joining is needed
if pd.notna(daf_final['join']).any():

  # get other children here
  child_rows = daf_final[pd.notna(daf_final['join'])]

  if any(child_rows['ID'].isin(child_rows['join'])):
    raise ValueError('Some of the join tables are related to eachother outside of their relationship with the parent row. Please fix this')
  

  for index, child_row in child_rows.iterrows():
    child_index = child_row['ID']
    parent_row = daf_final[daf_final['ID'].isin(child_row[['join']])]
    parent_index = parent_row.iloc[0]['ID']

    # check that the rows are idential
    parent_check = parent_row[['disaggregations','func','calculation','admin','q.type']].reset_index(drop=True)
    child_check = child_row.to_frame().transpose()[['disaggregations','func','calculation','admin','q.type']].reset_index(drop=True)

    check_result = child_check.equals(parent_check)

    if not check_result:
      raise ValueError('Joined rows are not identical in terms of admin, calculations, function and disaggregations')
    # get the data and dataframe indeces of parents and children
    child_tupple = [(i,tup) for i, tup in enumerate(test_new) if tup[1] == child_index]
    parent_tupple = [(i, tup) for i, tup in enumerate(test_new) if tup[1] == parent_index]

    child_tupple_data = child_tupple[0][1][0].copy()
    child_tupple_index = child_tupple[0][0]
    parent_tupple_data = parent_tupple[0][1][0].copy()
    parent_tupple_index = parent_tupple[0][0]

    # rename the data so that they are readable
    varnames = [parent_tupple_data['variable'][0],child_tupple_data['variable'][0]]
    dataframes =[parent_tupple_data, child_tupple_data]

    for var, dataframe in  zip(varnames, dataframes):
      rename_dict = {'mean': 'mean_'+var, 'count': 'count_'+var, 'perc': 'perc_'+var,
                     'min': 'min_'+var, 'max': 'max_'+var}

      for old_name, new_name in rename_dict.items():
        if old_name in dataframe.columns:
          dataframe.rename(columns={old_name: new_name},inplace=True)

    # get the lists of columns to keep and merge
    columns_to_merge = [item for item in parent_tupple_data.columns if not any(word in item for word in ['mean', 'count',
                                                                                                          'max','min',
                                                                                                          'perc','variable'])]
    columns_to_keep = columns_to_merge+ list(rename_dict.values())

    parent_tupple_data= parent_tupple_data.merge(
      child_tupple_data[child_tupple_data.columns.intersection(columns_to_keep)], 
      on = columns_to_merge,how='left')


    parent_index_f = parent_tupple[0][1][1]
    parent_label_f = str(child_tupple[0][1][2]).split()[0]+' & '+ str(parent_tupple[0][1][2])

    new_list = (parent_tupple_data,parent_index_f,parent_label_f)
    test_new[parent_tupple_index] = new_list

    del test_new[child_tupple_index]

# write excel files
print('Writing files')
filename = research_cycle+'_'+id_round+'_'+date

filename_dash = 'output/'+filename+'_dashboard.xlsx'
filename_toc = 'output/'+filename+'_TOC.xlsx'
filename_wide_toc = 'output/'+filename+'_wide_TOC.xlsx'

construct_result_table(test_new, filename_toc,make_pivot_with_strata = False)
construct_result_table(test, filename_wide_toc,make_pivot_with_strata = True)
concatenated_df.to_excel(filename_dash, index=False)
print('All done. congratulations')