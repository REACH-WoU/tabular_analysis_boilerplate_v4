import pandas as pd
import numpy as np
import re
from itertools import combinations
from openpyxl.styles import PatternFill, Font
from openpyxl import Workbook
import warnings
warnings.filterwarnings("ignore", 'This pattern is interpreted as a regular expression, and has match groups.')

# %load Functions.py
def load_tool_choices(filename_tool, label_colname, keep_cols=False):
     tool_choices = pd.read_excel(filename_tool, sheet_name="choices", dtype="str")

     if not keep_cols:
         tool_choices = tool_choices[['list_name', 'name', label_colname]]

     # Remove rows with missing values in 'list_name' column
     tool_choices = tool_choices.dropna(subset=['list_name'])

     # Keep only distinct rows
     tool_choices = tool_choices.drop_duplicates()

     # Convert to DataFrame
     tool_choices = pd.DataFrame(tool_choices)

     return(tool_choices)


def load_tool_survey(filename_tool, label_colname, keep_cols=False):
    tool_survey = pd.read_excel(filename_tool, sheet_name="survey", dtype="str")

    tool_survey = tool_survey.dropna(subset=['type'])

    tool_survey['q.type'] = tool_survey['type'].apply(lambda x: re.split(r'\s', x)[0])
    tool_survey['list_name'] = tool_survey['type'].apply(lambda x: re.split(r'\s', x)[1] if re.match(r'select_', x) else None)

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

    return tool_survey

def map_names(column_name,column_values_name,summary_table, tool_survey, tool_choices, na_include = False):
    choices_shortlist = tool_choices[
      tool_choices['list_name'].values==tool_survey[tool_survey['name']==summary_table[column_name][0]]['list_name'].values
      ][['name','label::English']]
    mapping_dict = dict(zip(choices_shortlist['name'], choices_shortlist['label::English']))
    if na_include is True:
      mapping_dict['No_data_available_NA'] = 'No data available (NA)'
    summary_table[column_values_name] = summary_table[column_values_name].map(mapping_dict)
    return summary_table

def weighted_mean(df, weight_column,numeric_column):
    weighted_sum = (df[numeric_column] * df[weight_column]).sum()
    total_weight = df[weight_column].sum()
    weighted_mean_result = weighted_sum / total_weight
    weighted_max_result = df[numeric_column].max()
    weighted_min_result = df[numeric_column].min()
    count = df.shape[0]
    return pd.Series({'mean': weighted_mean_result, 
                      'max': weighted_max_result,
                      'min': weighted_min_result,
                      'count': count})

def get_variable_type(data, variable_name):
    if data[variable_name].dtype == 'object':
        return 'string'
    elif data[variable_name].dtype == 'int64' or data[variable_name].dtype == 'int' or data[variable_name].dtype == 'int32':
        return 'integer'
    elif data[variable_name].dtype == 'float64' or data[variable_name].dtype == 'float' or data[variable_name].dtype == 'float32':
        return 'decimal'

def check_daf_filter(daf,data, filter_daf, tool_survey, tool_choices):
    merged_daf = filter_daf.merge(daf, on='ID', how='inner')
    # some calculate variables can be NaN
    merged_daf = merged_daf.drop(['calculation','join','disaggregations'], axis=1)
    # check if rows contain NaN
    if merged_daf.isnull().values.any():
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
                    raise ValueError(f"Operation {row['operation']} not allowed for string variables")
            continue

        if value_type == str:
            # check that the variable and the value have the same type
            if get_variable_type(data[row['datasheet']], row['variable_x']) != 'string':
                raise ValueError(f"Variable {row['variable_x']} has another type then filter value")
            # check that the operation is allowed for the type
            if row["operation"].strip(' ') not in ["!=", "=="]:
                raise ValueError(f"Operation {row['operation']} not allowed for string variables")
        else:
            # check that the variable and the value have the same type
            if get_variable_type(data[row['datasheet']], row['variable_x']) == 'string':
                raise ValueError(f"Variable {row['variable_x']} has another type then filter value")
            # check that the operation is allowed for the type
            if row["operation"].strip(' ') not in ["<", ">", "<=", ">=", "!=", "=="]:
                raise ValueError(f"Operation {row['operation']} not allowed for numeric variables")

def check_daf_consistency(daf, data, sheets, resolve=False):
    # check that all variables have a datasheet
    if daf['datasheet'].isnull().values.any():
        if not resolve:
            raise ValueError('the following are missing ' + ','.join(daf[daf['datasheet'].isnull().values]['variable']))
        else:
            print('the following are missing ' + ','.join(daf[daf['datasheet'].isnull().values]['variable']))
            daf.dropna(subset=['datasheet'], inplace=True)

    # check that all variables in daf are in the corresponding data sheets
    for id, row in daf.iterrows():
        if row["variable"] not in data[row["datasheet"]].columns:
            if not resolve:
                raise ValueError(f"Column {row['variable']} not found in {row['datasheet']}")
            else:
                print(f"Column {row['variable']} not found in {row['datasheet']}")
                daf.drop(id, inplace=True)

        if row["disaggregations"] not in ["overall", ""] and not pd.isna(row['disaggregations']):
            row["disaggregations"] = row["disaggregations"].replace(" ","")
            disaggregations_list = row["disaggregations"].split(",")

            for disaggregations_item in disaggregations_list:
                if disaggregations_item not in data[row["datasheet"]].columns:
                    error_message = f"Disaggregation {disaggregations_item} not found in {row['datasheet']} for variable {row['variable']}"
                    if not resolve:
                        raise ValueError(error_message)
                    else:
                        print(error_message)
                        daf.drop(id, inplace=True)
                        break

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
                raise ValueError(f"Intersection between {sheet1} and {sheet2} : {intersection}")
            else:
                print(f"Intersection between {sheet1} and {sheet2} : {intersection}")
                print("Resolve by removing from DAF the variables that are in both sheets")
                daf = daf[~daf['variable'].isin(intersection)]


    for sheet in sheets:
        # check that all sheets have variables in daf
        if not sheet_dict[sheet]:
            print(f"WARNING: Sheet {sheet} has no variables in DAF")
        print(f"Sheet {sheet} has {len(sheet_dict[sheet])} variables")

    return daf

def custom_sort_key(value):
    if value in 'Total' and isinstance(value, str):
        return 'zzzzzzzzzzz'  # This super dumb but it works
    else:
        return value

def make_pivot(table, index_list, column_list, value):
    pivot_table = table.pivot_table(index=index_list,
                                    columns=column_list,
                                    values=value).reset_index()
    return pivot_table


def construct_result_table(tables_list,file_name, make_pivot_with_strata = False):
    workbook = Workbook()
    workbook.create_sheet("Table_of_content", 0)
    workbook.create_sheet("Data", 1)
    content_sheet = workbook["Table_of_content"]
    data_sheet = workbook["Data"]
    link_idx = 1

    # add columns in the content sheet
    content_sheet.append(["ID", "Link"])

    for idx, element in enumerate(tables_list):
        table, ID, label = element
        values_variable = "perc" if "perc" in table.columns else "mean"
        if values_variable == "perc":
          if 'disaggregations_category_1' in table.columns:
            pivot_columns = ["disaggregations_category_1"]
          else:
            pivot_columns = []
          if "disaggregations_category_2" in table.columns:
              pivot_columns.append("disaggregations_category_2")

          if make_pivot_with_strata:
              if table['admin_category'].isin(['Total']).any():
                table_dirty  = table[table['admin_category']=='Total']
                table_clean  = table[table['admin_category']!='Total']

                pivot_table_dirty = make_pivot(table_dirty, pivot_columns + ["option"], ["admin_category"],values_variable)
                pivot_table_clean = make_pivot(table_clean, pivot_columns + ["option"], ["admin_category"], values_variable)

                pivot_table = pd.merge(pivot_table_clean,pivot_table_dirty[['option','Total']], on =['option'], how = 'left')
              else:
                pivot_table = make_pivot(table, pivot_columns + ["option"], ["admin_category"], values_variable)
          else:
              pivot_table = make_pivot(table, pivot_columns + ["admin_category", "count"], ["option"], values_variable)
              pivot_table = pivot_table.sort_values(by='admin_category', key=lambda x: x.map(custom_sort_key))

        else:
          pivot_table = table


        cell_id = f"A{link_idx}"
        link_idx += len(pivot_table) + 3
        data_sheet.append([label])
        data_sheet.append(list(pivot_table.columns))
        for _, row in pivot_table.iterrows():
            if values_variable == "perc":
              row_id = data_sheet.max_row + 1
              for i, value in enumerate(row):
                if isinstance(value, (float, np.float64, np.float32)) and not pd.isna(value):
                    cell = data_sheet.cell(row=row_id, column=i + 1)
                    cell.value = value
                    cell.number_format = '0.00%'
                else:
                  cell = data_sheet.cell(row=row_id, column=i + 1)
                  cell.value = value
        data_sheet.append([])

        text_on_link = label + ' ' + values_variable
        link_text = f'=HYPERLINK("#\'Data\'!{cell_id}", "{text_on_link}")'
        content_sheet.cell(row=idx + 2, column=2, value=link_text)
        content_sheet.cell(row=idx + 2, column=1, value=ID)

    for col_idx, column in enumerate(data_sheet.columns, 1):
        if col_idx == 1:
            data_sheet.column_dimensions[column[0].column_letter].width = 30
        else:
            data_sheet.column_dimensions[column[0].column_letter].width = 20

    content_sheet.column_dimensions['B'].width = 40
    for cell in content_sheet["B"][1:]:
        cell.font = Font(bold=True, color="FF0000FF")
    for cell in content_sheet["A"][1:]:
        cell.font = Font(bold=True)

    workbook.save(file_name)
    return workbook


def disaggregation_creator(daf_final, data, filter_dictionary,tool_choices, tool_survey, weight_column =None):

    if weight_column == None:
        for sheet in data:
            data[sheet]['weight']=1

    daf_final_freq = daf_final[daf_final['func'].isin(['freq', 'select_one', 'select_multiple'])]
    daf_final_num = daf_final[daf_final['func']=='numeric']


    daf_final_freq.reset_index(inplace=True)
    daf_final_num.reset_index(inplace=True)

    df_list = []

    if len(daf_final_freq)>0:
        for i, row in daf_final_freq.iterrows():
            # break down the disaggregations into a convenient list
            if not pd.isna(daf_final_freq.iloc[i]['disaggregations']):
              if ',' in daf_final_freq.iloc[i]['disaggregations']:
                  disaggregations = daf_final_freq.iloc[i]['disaggregations'].split(',')
                  disaggregations = [s.replace(" ", "") for s in disaggregations]
              else:
                  disaggregations = [daf_final_freq.iloc[i]['disaggregations']]
                  disaggregations = [s.replace(" ", "") for s in disaggregations]
            else:
              disaggregations = []
            if not pd.isna(daf_final_freq.iloc[i]['calculation']):
              # break down the calculations
              if ' ' in daf_final_freq.iloc[i]['calculation']:
                calc = daf_final_freq.iloc[i]['calculation'].split(',')
                calc = [x.strip(' ') for x in calc]
              else:
                calc = [daf_final_freq.iloc[i]['calculation']]
                calc = [x.strip(' ') for x in calc]
            else:
              calc = 'None'

            # get the correct sheet & add filters
            if daf_final_freq.iloc[i]['ID'] in filter_dictionary.keys():
              filter_text = 'data["'+daf_final_freq.iloc[i]['datasheet']+'"]['+filter_dictionary[daf_final_freq.iloc[i]['ID']]
              data_temp = eval(filter_text)
            else:
              data_temp  = data[daf_final_freq.iloc[i]['datasheet']]

            # keep only those columns that we'll need
            selected_columns = [daf_final_freq['variable'][i]]+disaggregations+[daf_final_freq['admin'][i]]+['weight']
            data_temp = data_temp[selected_columns]

            if 'include_na' in calc or 'add_total' in calc:
              data_temp.loc[:, daf_final_freq['variable'][i]] = data_temp[daf_final_freq['variable'][i]].fillna('No_data_available_NA')
              na_includer = True
            else:
              # remove NA rows
              data_temp = data_temp[data_temp[daf_final_freq['variable'][i]].notna()]
              na_includer = False

            if data_temp.shape[0]>0 :
            # keep a backup for select multiples
              data_temp_backup = data_temp.copy()

              # break down the data form SM
              if daf_final_freq.iloc[i]['q.type'] in ['select_multiple']:
                  data_temp.loc[:,daf_final_freq.iloc[i]['variable']] = data_temp[daf_final_freq.iloc[i]['variable']].str.strip()

                  data_temp.loc[:,daf_final_freq.iloc[i]['variable']] = data_temp[daf_final_freq.iloc[i]['variable']].str.split(' ').copy()
                  # Separate rows using explode
                  data_temp = data_temp.explode(daf_final_freq.iloc[i]['variable'], ignore_index=True)

              groupby_columns = [daf_final_freq['admin'][i]]+disaggregations+[daf_final_freq['variable'][i]]

              summary_stats=data_temp.groupby(groupby_columns)['weight'].agg(['sum'])
              # get the same stats but for the full subsample (not calculating option samples)
              groupby_columns_ov = [daf_final_freq['admin'][i]]+disaggregations

              summary_stats_var_om=data_temp_backup.groupby(groupby_columns_ov)['weight'].agg(['sum','count'])

              summary_stats.reset_index(inplace=True)
              summary_stats_var_om.reset_index(inplace=True)

              # rename them
              summary_stats.rename(columns = {'sum':'category_count'}, inplace=True)
              summary_stats_var_om.rename(columns = {'sum':'general_count'}, inplace=True)


              summary_stats_full = summary_stats.merge(summary_stats_var_om, on = groupby_columns_ov, how = 'left')


              new_column_names = {daf_final_freq['variable'][i]:'option',
                                  daf_final_freq['admin'][i]:'admin_category'}

              if disaggregations != []:
                for j, column_name in enumerate(disaggregations):
                    new_column_names[column_name] = f'disaggregations_category_{j+1}'

              summary_stats_full.rename(columns=new_column_names, inplace=True)


              summary_stats_full['admin'] = daf_final_freq['admin'][i]
              summary_stats_full['variable'] = daf_final_freq['variable'][i]

              if disaggregations != []:
                for j, column_name in enumerate(disaggregations):
                    summary_stats_full[f'disaggregations_{j+1}'] = disaggregations[j]


              # option replacer


              if tool_survey['name'].isin([daf_final.loc[i,'variable']]).any():
                summary_stats_full = map_names( column_name= 'variable',
                                              column_values_name='option',
                                                summary_table = summary_stats_full,
                                                tool_survey=tool_survey,
                                                tool_choices=tool_choices,
                                                na_include = na_includer)


              # disaggregations category replacer
              if disaggregations != [] and tool_survey['name'].isin(disaggregations).any():
                for j, column_name in enumerate(disaggregations):
                    if disaggregations[j] in set(tool_survey['name']):
                        summary_stats_full = map_names(column_name= f'disaggregations_{j+1}',column_values_name=f'disaggregations_category_{j+1}', summary_table = summary_stats_full,tool_survey=tool_survey,tool_choices=tool_choices)

              # admin category replacer

              if daf_final_freq['admin'].iloc[i] in set(tool_survey['name']):
                  summary_stats_full = map_names(column_name= 'admin',column_values_name='admin_category', summary_table = summary_stats_full,tool_survey=tool_survey,tool_choices=tool_choices)


              # add proper labels

              summary_stats_full['variable'] = daf_final_freq['variable_label'][i]
              if disaggregations != []:
                disaggregations_labels = daf_final_freq['disaggregations_label'][i]
                summary_stats_full[f'disaggregations_{j+1}'] = disaggregations_labels

              # add perc
              summary_stats_full['perc'] = round(summary_stats_full['category_count']/summary_stats_full['general_count'],4)

              summary_stats_full.drop(columns=['category_count', 'general_count'], inplace=True)


              if 'add_total' in calc:
                summary_stats_total=data_temp.groupby(daf_final_freq['variable'][i])['weight'].agg(['sum'])  #remove count here bruh
                summary_stats_total.reset_index(inplace=True)
                summary_stats_total['perc'] = round(summary_stats_total['sum']/data_temp_backup['weight'].sum(),4) # sometimes weights are wonky. so we're accounting for that
                summary_stats_total['count'] = data_temp_backup.shape[0] # add count (n of non-na rows)
                # drom the sum column
                summary_stats_total.drop(columns=['sum'], inplace=True)


                # rename columns
                new_column_names = {daf_final_freq['variable'][i]:'option'}
                summary_stats_total.rename(columns=new_column_names, inplace=True)
                # add new columns to match the existing format
                summary_stats_total['admin'] = 'Total'
                summary_stats_total['admin_category']= 'Total'
                summary_stats_total['variable'] = daf_final_freq['variable'][i]

                if tool_survey['name'].isin([daf_final.loc[i,'variable']]).any():
                  summary_stats_total = map_names(column_name= 'variable',
                                                  column_values_name='option',
                                                  summary_table = summary_stats_total,
                                                  tool_survey=tool_survey,
                                                  tool_choices=tool_choices,
                                                  na_include = na_includer)

                summary_stats_total['variable'] = daf_final_freq['variable_label'][i]
                if disaggregations != []:
                  for j, column_name in enumerate(disaggregations):
                    summary_stats_total[f'disaggregations_{j+1}'] = 'Total'
                    summary_stats_total[f'disaggregations_category_{j+1}'] = 'Total'
                
                summary_stats_full = pd.concat([summary_stats_full, summary_stats_total], ignore_index=True)


              if disaggregations != []:
                label = daf_final_freq.iloc[i]['variable']+' broken down by '+ daf_final_freq.iloc[i]['disaggregations_label'] + ' on the admin of '+daf_final_freq.iloc[i]['admin']
              else:
                label = daf_final_freq.iloc[i]['variable']+' on the admin of '+daf_final_freq.iloc[i]['admin']

              disagg_columns = [col for col in summary_stats_full.columns if col.startswith('disaggregations')]
              summary_stats_full['ID'] = daf_final_freq.iloc[i]['ID']
              columns = ['ID','admin','admin_category','option','variable']+ disagg_columns + ['perc','count']
              summary_stats_full = summary_stats_full[columns]
              df_list.append((summary_stats_full, daf_final_freq['ID'][i], label))


    if len(daf_final_num)>0:
    # Deal with numerics
        for i, row in daf_final_num.iterrows():
          if not pd.isna(daf_final_num.iloc[i]['disaggregations']):
              if ' ' in daf_final_num.iloc[i]['disaggregations']:
                  disaggregations = daf_final_num.iloc[i]['disaggregations'].split(' ')
              else:
                  disaggregations = [daf_final_num.iloc[i]['disaggregations']]
          else:
            disaggregations = []


          # get the correct sheet & add filters
          if daf_final_num.iloc[i]['ID'] in filter_dictionary.keys():
            filter_text = 'data["'+daf_final_num.iloc[i]['datasheet']+'"]['+filter_dictionary[daf_final_num.iloc[i]['ID']]
            data_temp = eval(filter_text)
          else:
            data_temp  = data[daf_final_num.iloc[i]['datasheet']]

          # keep only those columns that we'll need
          selected_columns = [daf_final_num['variable'][i]]+disaggregations+[daf_final_num['admin'][i]]+['weight']
          data_temp = data_temp[selected_columns]

          # drop all NA observations
          data_temp = data_temp[data_temp[daf_final_num['variable'][i]].notna()]

          if data_temp.shape[0]>0 :

            groupby_columns = disaggregations+[daf_final_num['admin'][i]]

            # get the general disaggregations statistics

            summary_stats = data_temp.groupby(groupby_columns).apply(weighted_mean, weight_column='weight',numeric_column =daf_final_num['variable'][i])

            summary_stats= summary_stats.reset_index()

            new_column_names = {daf_final_num['admin'][i]:'admin_category'}

            if disaggregations != []:
              for j, column_name in enumerate(disaggregations):
                  new_column_names[column_name] = f'disaggregations_category_{j+1}'



            summary_stats.rename(columns=new_column_names, inplace=True)

            summary_stats['admin'] = daf_final_num['admin'][i]
            summary_stats['variable'] = daf_final_num['variable'][i]

            if disaggregations != []:
              for j, column_name in enumerate(disaggregations):
                  summary_stats[f'disaggregations_{j+1}'] = disaggregations[j]

            # disaggregations category replacer
            if disaggregations != [] and tool_survey['name'].isin(disaggregations).any():
              for j, column_name in enumerate(disaggregations):
                  if disaggregations[j] in set(tool_survey['name']):
                      summary_stats = map_names(column_name= f'disaggregations_{j+1}',column_values_name=f'disaggregations_category_{j+1}', summary_table = summary_stats,tool_survey=tool_survey,tool_choices=tool_choices)

            # admin category replacer
            if daf_final_num['admin'].iloc[i] in set(tool_survey['name']):
                summary_stats = map_names(column_name= 'admin',column_values_name='admin_category', summary_table = summary_stats,tool_survey=tool_survey,tool_choices=tool_choices)


            # add proper labels
            summary_stats['variable'] = daf_final_num['variable_label'][i]
            if disaggregations != []:
              disaggregations_labels = daf_final_num['disaggregations_label'][i]
              summary_stats[f'disaggregations_{j+1}'] = disaggregations_labels

            if 'add_total' in calc:
              summary_stats_total = weighted_mean(data_temp,weight_column='weight',numeric_column =daf_final_num['variable'][i]).to_frame().transpose()

              # add new columns to match the existing format
              summary_stats_total['admin'] = 'Total'
              summary_stats_total['admin_category']= 'Total'
              summary_stats_total['variable'] = daf_final_num['variable'][i]

              summary_stats_total['variable'] = daf_final_num['variable_label'][i]
              if disaggregations != []:
                for j, column_name in enumerate(disaggregations):
                  summary_stats_total[f'disaggregations_{j+1}'] = 'Total'
                  summary_stats_total[f'disaggregations_category_{j+1}'] = 'Total'

              summary_stats = pd.concat([summary_stats, summary_stats_total], ignore_index=True)
            if disaggregations != []:
              label = daf_final_num.iloc[i]['variable']+' broken down by '+ daf_final_num.iloc[i]['disaggregations_label'] + ' on the admin of '+daf_final_num.iloc[i]['admin']
            else:
              label = daf_final_num.iloc[i]['variable']+' on the admin of '+daf_final_num.iloc[i]['admin']

            disagg_columns = [col for col in summary_stats_full.columns if col.startswith('disaggregations')]
            summary_stats['ID'] = daf_final_num.iloc[i]['ID']
            columns = ['ID','admin','admin_category','variable']+disagg_columns + ['mean','min','max','count']
            summary_stats = summary_stats[columns]

            df_list.append((summary_stats, daf_final_num['ID'][i], label))
    return(df_list)
