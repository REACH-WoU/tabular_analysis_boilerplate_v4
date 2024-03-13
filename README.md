# Tabular analysis V4  
This boilerplate is an improved version of [tabular_analysis_boilerplate_v3](https://github.com/REACH-WoU/tabular_analysis_boilerplate_v3) with similar use functionality and requirements. The main difference between the two versions is that V4 is written in Python as opposed to R and thus, is much faster. The text below guides the user how to use the script to run different analyses. This script produces a traditional TOC table, a wide TOC table that uses the admin variable to pivot the tables as well as a long table that serves as dashboard inputs for dashboards that work on frequency tables.

## Table of Contents
- [Basic inputs](#Basic-inputs-to-the-script)
- [Filling out the DAF form](#Filling-out-the-DAF-form)
- [Running the script](#Running-the-script)
  - [Inputs](#Inputs)
  - [Checks](#Checks)
- [Outputs](#outputs)


# Basic inputs to the script
As the previous script V4 requires the user to input the kobo tool and DAF form into the `resources` folder of the V4 main folder. The user will also need to input their dataframe into the `data` folder of the V4. For users who are not comfortable with writing Python scripts it is recomended that the uploaded dataframe is the final clean version of the data that contains all of the variables that the user wants to analyse. If the user will need to create additional variables, the V4 script has a small subsection that provides the user with the space to do so.

# Filling out the DAF form
The general way of filling out the DAF form is similar between V4 and V3, main differences between V3 and V4 will be marked in **bold**. The main difference between the structure of the excel DAF forms is that V4 DAF form has two sheets - `main` - same sheet as the V3 DAF sheet and `filter` - a sheet dedicated to the filtering functionality of the V4 script.
The DAF form example is provided in the `resources/DAF_example.xlsx`. The basic structure of the file is the following:
|ID|	variable|	variable_label|	calculation|	func|	admin|	disaggregations	|disaggregations_label|	join|
|--|----------|---------------|------------|------|------|------------------|---------------------|-----|
|**Row index, unique to each row and should start with 1**| The name of the variable, should match exactly what you have in your Kobo tool and your dataframe, any differences will produce errors in the script|**The label of your variable, what did you ask the respondent?**| supports two functions `include_na` and `add_total`, see more below. | whether the variable should be disaggregated as a frequency  or as a weighted mean | The admin unit to be used for the disaggregation| What is the disaggregation variable you want to use for your `variable`?| **A nice label of your disaggregation column**| The `ID` of the parent row of the dependent table|

Some details for relevant columns:
- `ID` - new column, please fill it in. **Each row has to be unique**
- `calculation` - If you want to run a frequency analysis specify `freq`, `select_one`, `select_multiple` in the cell. If you want to get a weighted mean for the variable, specify `numeric` in the cell
- `func` - The current script supports two specifications of this columns, the same as V3. `include_na` Replaces NA values of the `variable` with `No data available (NA)`. `add_total` does the same but adds the general frequency table of the dependent variabme ommiting all of the entries in `admin` and `disaggregations` columns. You can leave it blank if you don't care about any of this.
- `disaggregations` the current version of the scripts supports multiple disaggregation columns, to use this, enter multiple disaggregation names in the cell and separate them with a comma `,`
- `admin` - works the same as in the previous version if you want to get the overall value, input `Overall`
- `join` - for cases where you want to make a table wider by merging a few tables together, input the `ID` of the parent table into the cell of `join` column. It is recommended that you only use this functionality if you have the same values in `disaggregations`,`func`,`calculation`,`admin`,`q.type` of the different rows. It is also required that the variables are related only with relationships of type `parent`-`child`. This essentially means that if you have the following table:

|ID|	variable|	variable_label|	calculation|	func|	admin|	disaggregations	|disaggregations_label|	join|
|--|----------|---------------|------------|------|------|------------------|---------------------|-----|
|1| variable_1|variable_1_labl|numeric| |Overall | age_group|Age group||
|2| variable_2|variable_2_labl|numeric| |Overall | age_group|Age group|1|
|3| variable_3|variable_3_labl|numeric| |Overall | age_group|Age group|1|

In this example a `parent`-`child` relationships mean 'variable_1' is the `parent` table, while `variable_2` and `variable_3` are the `child` tables. This table will create a numeric breakdown table that will have mean, min, max and count columns for all 3 variables. The individual tables for each variables will be removed from the output files. However, if you've input this relationship as:
|ID|	variable|	variable_label|	calculation|	func|	admin|	disaggregations	|disaggregations_label|	join|
|--|----------|---------------|------------|------|------|------------------|---------------------|-----|
|1| variable_1|variable_1_labl|numeric| |Overall | age_group|Age group||
|2| variable_2|variable_2_labl|numeric| |Overall | age_group|Age group|1|
|3| variable_3|variable_3_labl|numeric| |Overall | age_group|Age group|2|

This wouldn't work, as now, we're talking about a `grandparent`-`parent`-`child` relationship. Those aren't supported by the current script.

To add filters please fill in the table on the `filter` sheet of your DAF

|ID|	variable|	operation|	value|
|--|----------|---------------|------------|
|The id of the `variable` in the `main` sheet| The filtering variable|Filtering operation|Filter|

An example of the use of this functionality would be the following

|ID|	variable|	operation|	value|
|--|----------|---------------|------------|
|2| Age|>|18|

Meaning that the disaggregation presented on the `main` sheet with `ID` 2 will be calculated only for cases where the `Age` variable is greater than 18.
The `operation` column supports the following operations "<", ">", "<=", ">=", "!=", "==". You can filter your `main` disaggregation with 3 types of filtering operations:
 - Numeric filter (e.g. variable > 5)
 - Character filter (e.g. variable == Yes) **No quotation marks are needed**
 - Variable filter (e.g. variable > variable2) **Be careful when using this**

## Inputs

Prior to running the script please fill in:
 - Your working directory in row 6
 - The name of your research cycle in row 14
 - The round of your research cycle in row 15
 - The relevant relative paths and names of your Data, Kobo tool and DAF files in rows 18-20
 - The name of your label column **Must be identical in Kobo tool and survey sheets!!**
 - If you want to add any new variables not present in the data, please do so in the block that starts on line 34, please note that the new variables will be assigned a frequency type of `select_one` if they are not present in the Kobo tool.

## Checks
The script goes through the following checks:
- Check if all the mentioned variables are present in the data
- Check if all variables in `disaggregations`,`admin` are present in the datasheet where the `variable` is located
- Check if `variable` is present on more than 1 datasheet
- Check if the filter table was properly filled

# Outputs

The script will produce the following tables:
- TOC table - identical to the V3's TOC table
- TOC table wide - Same as the above table but pivoted by the admin variable to produce wider tables for more geographically inclined DAF files
- Dashboard input table - a table that is designed to be a better fit for dashboards that require a multivariate frequency table input








