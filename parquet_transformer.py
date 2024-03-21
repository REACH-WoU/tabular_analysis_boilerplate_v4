import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

# enter that path do your data
excel_path_data = 'data/test_frame.xlsx'

data = pd.read_excel(excel_path_data, sheet_name=None)
sheets = list(data.keys())

for id, dataframe in enumerate(data):
    table = pa.Table.from_pandas(data[sheets[id]])
    filename = 'data/parquet_inputs/'+sheets[id]+'.parquet'
    pq.write_table(table, filename, compression='BROTLI')
