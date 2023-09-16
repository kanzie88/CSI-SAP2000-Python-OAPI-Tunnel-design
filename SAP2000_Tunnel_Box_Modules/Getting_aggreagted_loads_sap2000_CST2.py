# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 13:03:00 2023

@author: 10022767
"""

import xlwings as xw
import pandas as pd
import numpy as np




#%% MOMENT AND AXIAL DESIGN

def MOMENT_AND_AXIAL_DESIGN(section_type,df,frames_LS_df_moment,frames_LS_df_shear):
   
    
    def resultcheck(Frame,OutputCase):
        
        P_M=result.loc[(result['Frame']==Frame) & (result['OutputCase']==OutputCase)][["P","M3"]]
        
        
        return [P_M["P"].values[0],P_M["M3"].values[0]]
    
    
    df=df.astype({"OutputCase":'str'}).reset_index(drop=True)
    df[["P","M3"]]=df[["P","M3"]].apply(pd.to_numeric, errors='coerce')
    
    
    # create dfs for ULS
    df_ULS = df[df['OutputCase'].str.startswith('1')]
    
    df_ULS["M3-abs"]=df_ULS["M3"].abs()
    df_ULS.sort_values(by=["M3-abs"],inplace=True,ascending=False)
    
    
    # aggregating and generating the 4 load_scenarios
    maxMz_index = df_ULS.groupby(['Frame'])['M3'].apply(lambda x: x.abs().idxmax())
    minMz_index = df_ULS.groupby(['Frame'])['M3'].apply(lambda x: x.abs().idxmin())
    maxFx_index = df_ULS.groupby(['Frame'])['P'].apply(lambda x: x.abs().idxmax())
    minFx_index = df_ULS.groupby(['Frame'])['P'].apply(lambda x: x.abs().idxmin())
    
    columns_to_keep=[
        'Frame',
        'OutputCase',
        'Station',
        'P',
        'M3'
      ]
    
    maxMz_df=df_ULS.loc[maxMz_index][columns_to_keep]
    maxMz_df["LS"]='Max Mz'
    minMz_df=df_ULS.loc[minMz_index][columns_to_keep]
    minMz_df["LS"]='Min Mz'
    maxFx_df=df_ULS.loc[maxFx_index][columns_to_keep]
    maxFx_df["LS"]='Max Fx'
    minFx_df=df_ULS.loc[minFx_index][columns_to_keep]
    minFx_df["LS"]='Min Fx'
    
    list_of_LSdf=[maxMz_df,minMz_df,maxFx_df,minFx_df]
    
    
    # prepare a compilation dataframe to be used
    frames_LS_df_compiled=frames_LS_df_moment.copy(deep=True)
    columns_to_add=['LS','P','M3','OutputCase','Station']
    for i in columns_to_add:
        frames_LS_df_compiled.insert(loc=2,column=i,value=np.nan)
    frames_LS_df_compiled=frames_LS_df_compiled.iloc[0:0]
    
    # # Append this compilation datframe it with the 4 load_scnearios df and the specific frame
    for LSdf in list_of_LSdf:
        # frames_LS_df_compiled=frames_LS_df_compiled.append(frames_LS_df_moment.merge(LSdf,on="Frame"))
        frames_LS_df_compiled=pd.concat([frames_LS_df_compiled,frames_LS_df_moment.merge(LSdf,on="Frame")])
        
    
    frames_LS_df_compiled.sort_values(by=["Frame","LS"],inplace=True)

    
    
    
    # creating a index for lookup
    frames_LS_df_compiled["Combined"]=frames_LS_df_compiled["Frame"]+frames_LS_df_compiled["LS"]
    # Extract the 'Column3' data
    column_data = frames_LS_df_compiled["Combined"]
    # Remove 'Column3' from the DataFrame
    frames_LS_df_compiled.drop(columns='Combined', inplace=True)
    # Insert 'Column3' at index 1
    frames_LS_df_compiled.insert(loc=0, column='Combined', value=column_data)
    
    # for the identified LS to find the corresponding SLS load combi index and extract out the loads
    frames_LS_df_compiled_SLS=frames_LS_df_compiled[["Frame",'Location','OutputCase',"LS",'Station','Combined']]
    frames_LS_df_compiled_SLS["OutputCase"]="2"+frames_LS_df_compiled_SLS["OutputCase"].str[1:]
    frames_LS_df_compiled_SLS=frames_LS_df_compiled_SLS.merge(df,on=["Frame",'OutputCase','Station'])[['Combined', 'Frame', 'Location', 'Station', 'OutputCase', 'M3', 'P',
           'LS']]
    
    frames_LS_df_compiled['OutputCase']=frames_LS_df_compiled['OutputCase'].str[:3]
    frames_LS_df_compiled_SLS['OutputCase']=frames_LS_df_compiled_SLS['OutputCase'].str[:3]
    
    
    
# =============================================================================
#     SHEAR DESIGN
# =============================================================================
    
    # extracting forces from SAP2000 output
    df=df.astype({"OutputCase":'str'}).reset_index(drop=True)
    df[["V2"]]=df[["V2"]].apply(pd.to_numeric, errors='coerce')
    
    # create dfs for ULS
    df_ULS = df[df['OutputCase'].str.startswith('1')]
    
    columns_to_keep=[
        'Frame',
        'OutputCase',
        'Station',
        'P',
        'M3',
        'V2'
      ]
    
    
    
    # prepare a compilation dataframe to be used
    frames_df_compiled=frames_LS_df_shear.copy(deep=True)
    columns_to_add=['V2','OutputCase','Station']
    for i in columns_to_add:
        frames_df_compiled.insert(loc=2,column=i,value=np.nan)
    frames_df_compiled=frames_df_compiled.iloc[0:0]
    
    # aggregating and generating the 4 load_scenarios
    maxV2_index = df_ULS.groupby(['Frame'])['V2'].apply(lambda x: x.abs().idxmax())
    
    maxV2_df=df_ULS.loc[maxV2_index][columns_to_keep]
    maxV2_df["LS"]='Max V2'
    
    # # Append this compilation datframe it with the 4 load_scnearios df and the specific frame
    frames_df_compiled=frames_LS_df_shear.merge(maxV2_df,on="Frame")
    
    frames_df_compiled.sort_values(by=["Frame"],inplace=True)
    
    # creating a index for lookup
    frames_df_compiled["Combined"]=frames_df_compiled["Frame"]+frames_df_compiled["LS"]
    # Extract the 'Column3' data
    column_data = frames_df_compiled["Combined"]
    # Remove 'Column3' from the DataFrame
    frames_df_compiled.drop(columns='Combined', inplace=True)
    # Insert 'Column3' at index 1
    frames_df_compiled.insert(loc=0, column='Combined', value=column_data)
    
    # Reorder 'Station' data
    column_data = frames_df_compiled["Station"]
    # Remove 'Station' from the DataFrame
    frames_df_compiled.drop(columns='Station', inplace=True)
    # Insert 'Station' at index 3
    frames_df_compiled.insert(loc=3, column='Station', value=column_data)
    
    # for the identified LS to find the corresponding SLS load combi index and extract out the loads
    frames_df_compiled_SLS=frames_df_compiled[["Frame",'Location','OutputCase',"LS",'Station','Combined']]
    frames_df_compiled_SLS["OutputCase"]="2"+frames_df_compiled["OutputCase"].str[-2:]
    frames_df_compiled_SLS=frames_df_compiled_SLS.merge(df,on=["Frame",'OutputCase','Station'])[['Combined', 'Frame', 'Location', 'Station', 'OutputCase','P','M3',"V2",'LS']]
    frames_df_compiled_SLS.drop_duplicates(inplace=True)
    
    return frames_LS_df_compiled,frames_LS_df_compiled_SLS,frames_df_compiled,frames_df_compiled_SLS



def populating_moment(Design_sheet_path,frames_LS_df_compiled,frames_LS_df_compiled_SLS):
    
    # Populate it into design spreadsheet
    wb = xw.Book(Design_sheet_path)
    
    # input into ULS sheet
    ws = wb.sheets["Summary-output-ULS"]
    ws["A1:G1000"].clear()
    print("previous ULS loads cleared")
    ws["A1"].options(pd.DataFrame,headers=True, index=False, expand='table').value = frames_LS_df_compiled
    print("ULS loads populated")
    
    # input into SLS sheet
    ws = wb.sheets["Summary-output-SLS"]
    ws["A1:G1000"].clear()
    print("previous SLS loads cleared")
    ws["A1"].options(pd.DataFrame,headers=True, index=False, expand='table').value = frames_LS_df_compiled_SLS
    print("SLS loads populated")
    
    
    # ws = wb.sheets["Wall"]
    # # ws["G10"].value="SECTION "+section_type

def populating_shear(Design_sheet_path,frames_df_compiled,frames_df_compiled_SLS):
    # Populate it into design spreadsheet
    wb = xw.Book(Design_sheet_path)
    
    
    # input into ULS sheet
    ws = wb.sheets["Summary-output-ULS"]
    ws["A1:G1000"].clear()
    print("previous ULS loads cleared")
    ws["A1"].options(pd.DataFrame,headers=True, index=False, expand='table').value = frames_df_compiled
    print("ULS loads populated")
    
    # input into SLS sheet
    ws = wb.sheets["Summary-output-SLS"]
    ws["A1:G1000"].clear()
    print("previous SLS loads cleared")
    ws["A1"].options(pd.DataFrame,headers=True, index=False, expand='table').value = frames_df_compiled_SLS
    print("SLS loads populated")
    
    # ws = wb.sheets["Section-All"]
    # ws["C10"].value="SECTION "+section_type


