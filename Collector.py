import pandas as pd
import os
import datetime as dt
from Operator import Operator
import time

# ROOT_FILE = "pokemon_data"
# COLUMN_INDEX = "Name"
# SPLIT_PARAMETERS = ["Type 1","Type 2"]
# ROOT_FILE = "Aging as of December 6th - weekly"
ROOT_FILE = "1"
COLUMN_INDEX = "Billing Document"
SPLIT_PARAMETERS = ["Country","Collector"]

start_time = time.time()

now = dt.datetime.now()
curent_time = now.strftime("%Y_%m_%d_%H_%M")
files_for_update = os.listdir("./Files for update")
df_root = pd.read_excel(f"{ROOT_FILE}.xlsx",index_col=COLUMN_INDEX,header=2)
df_header_list = df_root.columns.to_list()
df_root.to_excel(f"./History/{ROOT_FILE}{curent_time}.xlsx",
                                    index=False)
new_main_list = []
new_sub_list = []
new_junior_list = []
for file in files_for_update:
    df_update = pd.read_excel(f"./Files for update/{file}",index_col=COLUMN_INDEX,header=2)
    new_main_list.extend(df_update[SPLIT_PARAMETERS[0]].unique())
    new_sub_list.extend(df_update[SPLIT_PARAMETERS[1]].unique())
    # new_junior_list.extend(df_update[SPLIT_PARAMETERS[2]].unique())
    df_root.update(df_update)
df_root=df_root.reset_index()
df_root=df_root.reindex(df_header_list,axis=1)
# df_root.set_index("#", inplace=True)
# df_root.reset_index(inplace=True)
name_for_update = f"{ROOT_FILE}_{curent_time}"
df_root.to_excel(f"./Export/{name_for_update}.xlsx",startrow=2,
                                    index=False)
time.sleep(5)
operator = Operator(f"./Export/{name_for_update}")
operator.sorting_list = SPLIT_PARAMETERS
operator.main_list = new_main_list
operator.sub_list = new_sub_list
operator.spilt()
#
print("Time to update --- %s seconds ---" % (time.time() - start_time))