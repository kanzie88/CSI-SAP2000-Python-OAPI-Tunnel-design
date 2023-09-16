# -*- coding: utf-8 -*-
"""
Created on Fri Aug  4 06:41:30 2023

@author: 10022767
"""

import pandas as pd
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import numpy as np

def extract_pile_loads(df_loadcases,df_loadcombi,Load_summary_df,df_joint_coords_piles):
    
    Load_summary_df.iloc[:,2:21]=Load_summary_df.iloc[:,2:21].astype('float')
    df_loadcases=df_loadcases.astype({"F3":'int','OutputCase':'str'})
    df_loadcombi=df_loadcombi.astype({"F3":'int','OutputCase':'str'})
    # only aggregating the serv load combi
    df_loadcombi=df_loadcombi.loc[df_loadcombi.OutputCase.str.startswith("1") | (df_loadcombi.OutputCase=='UPL_ULS')] 
    subset=['Joint', 'Elm', 'OutputCase', 'F1', 'F2', 'F3','M1', 'M2', 'M3']
    df_loadcombi.drop_duplicates(subset=subset,inplace=True)
    df_loadcases.drop_duplicates(subset=subset,inplace=True)
    internal_pile_joint=df_joint_coords_piles[df_joint_coords_piles["Outer/Inner"]=="Inner"].Joint.iloc[0]
    outer_pile_joint=list(df_joint_coords_piles[df_joint_coords_piles["Outer/Inner"]=="Outer"].Joint)
    
    # Generating the max pile loads in DL and LL
    # =============================================================================
    #     # extract the joint loads for the pile locations for innerpiles        
    # =============================================================================
    df_loadcombi_inner=df_loadcombi.loc[df_loadcombi.Joint==internal_pile_joint]

# =============================================================================
#     # extract the joint loads for the pile locations for outterpiles        
# =============================================================================
    df_loadcombi_outer=df_loadcombi.loc[ df_loadcombi.Joint.isin(outer_pile_joint)]
    
    
    #(a) produce df_temp Load_combination=="103" for temp condition
    max_temp_comp_case="103"
    df_loadcombi_temp=df_loadcombi_outer.loc[df_loadcombi_outer["OutputCase"]==max_temp_comp_case].reset_index(drop=True)
    # find Outputcase with max F3 for df_temp
    max_temp_comp_index = df_loadcombi_temp.F3.idxmax()
    max_temp_comp_df=df_loadcombi_temp.loc[max_temp_comp_index]
    max_temp_comp=max_temp_comp_df.F3
    max_temp_comp_case=str(max_temp_comp_df.OutputCase)
    max_temp_comp_joint=max_temp_comp_df.Joint
    
    
    
    # (b)produce df_perm without Load_combination=="103" for perm condition
    df_loadcombi_perm=df_loadcombi_outer.loc[df_loadcombi_outer["OutputCase"]!="103"].reset_index(drop=True)
    # find Outputcase with max F3 for df_perm
    maxZ_perm_index = df_loadcombi_perm.F3.idxmax()
    max_perm_comp_df=df_loadcombi_perm.loc[maxZ_perm_index]
    max_perm_comp=max_perm_comp_df.F3
    max_perm_comp_case=str(max_perm_comp_df.OutputCase)
    max_perm_comp_joint=max_perm_comp_df.Joint
    
    
    
    # (c)Check for tension
    minZ_perm_index = df_loadcombi_perm.F3.idxmin()
    max_perm_tension_df=df_loadcombi_perm.loc[minZ_perm_index]
    max_perm_tension=max_perm_tension_df.F3
    max_perm_tension_case=str(max_perm_tension_df.OutputCase)
    max_perm_tension_joint=max_perm_tension_df.Joint
    
    
    
    columns_to_keep=['load_case','Type']
    
    # (a)
    # filter the  Load_summary_df column to just the load combi of concern
    # max_temp_comp_case
    columns_to_keep.append(max_temp_comp_case)
    Load_summary_df_temp=Load_summary_df.loc[:,columns_to_keep]
    
    # merge the Load_summary_df column with the df_loadcases
    df_loadcases.OutputCase=df_loadcases.OutputCase.astype("str")
    
    # filter out load cases prior to merge
    df_loadcases_a=df_loadcases.loc[(df_loadcases.Joint==max_temp_comp_joint)]
    
    Load_summary_df_temp=Load_summary_df_temp.merge(df_loadcases_a,left_on='load_case',right_on='OutputCase')
    
    Load_summary_df_temp=Load_summary_df_temp[Load_summary_df_temp[max_temp_comp_case]!=0]
    
    Load_summary_df_temp["unfactored-per m run"]=Load_summary_df_temp.F3
    
    Load_summary_df_temp["CASETYPE"]='max_temp_comp'
    Load_summary_df_temp["load_case"]=max_temp_comp_case
    
    #%%
    # load cases to be considered
    columns_to_keep=['load_case','Type']
    # (b)
    # filter the  Load_summary_df column to just the load combi of concern
    # max_perm_comp_case
    columns_to_keep.append(max_perm_comp_case)
    Load_summary_df_comp=Load_summary_df.loc[:,columns_to_keep]
    
    # merge the Load_summary_df column with the df_loadcases
    df_loadcases.OutputCase=df_loadcases.OutputCase.astype("str")
    
    # filter out load cases prior to merge
    df_loadcases_b=df_loadcases.loc[(df_loadcases.Joint==max_perm_comp_joint)]
    
    Load_summary_df_comp=Load_summary_df_comp.merge(df_loadcases_b,left_on='load_case',right_on='OutputCase')
    Load_summary_df_comp=Load_summary_df_comp[Load_summary_df_comp[max_perm_comp_case]!=0]
    
    Load_summary_df_comp["unfactored-per m run"]=Load_summary_df_comp.F3
    
    Load_summary_df_comp["CASETYPE"]='max_perm_comp'
    Load_summary_df_comp["load_case"]=max_perm_comp_case
    
    #%%
    # load cases to be considered
    columns_to_keep=['load_case','Type']
    # (c)
    # filter the  Load_summary_df column to just the load combi of concern
    # max_tension_comp_
    columns_to_keep.append(max_perm_tension_case)
    Load_summary_df_tension=Load_summary_df.loc[:,columns_to_keep]
    Load_summary_df_tension[max_perm_tension_case]
    
    
    # merge the Load_summary_df column with the df_loadcases
    df_loadcases.OutputCase=df_loadcases.OutputCase.astype("str")
    
    
    # filter out load cases prior to merge
    df_loadcases_c=df_loadcases.loc[(df_loadcases.Joint==max_perm_tension_joint)]
    
    Load_summary_df_tension=Load_summary_df_tension.merge(df_loadcases_c,left_on='load_case',right_on='OutputCase')
    Load_summary_df_tension=Load_summary_df_tension[Load_summary_df_tension[max_perm_tension_case]!=0]
    
    Load_summary_df_tension["unfactored-per m run"]=Load_summary_df_tension.F3
    
    Load_summary_df_tension["CASETYPE"]='max_tension'
    Load_summary_df_tension["load_case"]=max_perm_tension_case
    

    
   
    
    
    
    #%%

    Load_summary_all=pd.concat([Load_summary_df_temp,Load_summary_df_comp,Load_summary_df_tension],axis=0)
    final_outer=Load_summary_all.groupby(by=["CASETYPE","Type","load_case"]).agg({"unfactored-per m run":sum}).reset_index()
    final_outer.load_case.replace({"UPL_ULS":"UPL_SLS"},inplace=True)
    
    # add in UPL_ULS load
    a={"CASETYPE":["max_tension_ULS"],"Type":["DL"],"load_case":[max_perm_tension_case],"unfactored-per m run":[df_loadcombi_outer[df_loadcombi_outer["OutputCase"]=="UPL_ULS"].F3.min()]}
    UPL_ULS=pd.DataFrame(a)
    final_outer=pd.concat([final_outer,UPL_ULS],axis=0)
    
    
    #%%
# =============================================================================
# for Innerpiles
# =============================================================================
    
    if internal_pile_joint is np.nan:
        return final_outer,"No Inner Piles"
    else:
        
       
            
    
      
        
        #(a) produce df_temp Load_combination=="103" for temp condition
        max_temp_comp_case="103"
        df_loadcombi_temp=df_loadcombi_inner.loc[df_loadcombi_inner["OutputCase"]==max_temp_comp_case].reset_index(drop=True)
        # find Outputcase with max F3 for df_temp
        max_temp_comp_index = df_loadcombi_temp.F3.idxmax()
        max_temp_comp_df=df_loadcombi_temp.loc[max_temp_comp_index]
        max_temp_comp=max_temp_comp_df.F3
        max_temp_comp_case=str(max_temp_comp_df.OutputCase)
        max_temp_comp_joint=max_temp_comp_df.Joint
        
        
        
        # (b)produce df_perm without Load_combination=="103" for perm condition
        df_loadcombi_perm=df_loadcombi_inner.loc[df_loadcombi_inner["OutputCase"]!="103"].reset_index(drop=True)
        # find Outputcase with max F3 for df_perm
        maxZ_perm_index = df_loadcombi_perm.F3.idxmax()
        max_perm_comp_df=df_loadcombi_perm.loc[maxZ_perm_index]
        max_perm_comp=max_perm_comp_df.F3
        max_perm_comp_case=str(max_perm_comp_df.OutputCase)
        max_perm_comp_joint=max_perm_comp_df.Joint
        
        
        
        # (c)Check for tension
        minZ_perm_index = df_loadcombi_perm.F3.idxmin()
        max_perm_tension_df=df_loadcombi_perm.loc[minZ_perm_index]
        max_perm_tension=max_perm_tension_df.F3
        max_perm_tension_case=str(max_perm_tension_df.OutputCase)
        max_perm_tension_joint=max_perm_tension_df.Joint
        
        
        
        
        # load cases to be considered
        columns_to_keep=['load_case','Type']
        
        # (a)
        # filter the  Load_summary_df column to just the load combi of concern
        # max_temp_comp_case
        columns_to_keep.append(max_temp_comp_case)
        Load_summary_df_temp=Load_summary_df.loc[:,columns_to_keep]
        
        # merge the Load_summary_df column with the df_loadcases
        df_loadcases.OutputCase=df_loadcases.OutputCase.astype("str")
        
        # filter out load cases prior to merge
        df_loadcases_a=df_loadcases.loc[(df_loadcases.Joint==max_temp_comp_joint)]
        
        Load_summary_df_temp=Load_summary_df_temp.merge(df_loadcases_a,left_on='load_case',right_on='OutputCase')
        Load_summary_df_temp=Load_summary_df_temp[Load_summary_df_temp[max_temp_comp_case]!=0]
        
        Load_summary_df_temp["unfactored-per m run"]=Load_summary_df_temp.F3
        
        Load_summary_df_temp["CASETYPE"]='max_temp_comp'
        Load_summary_df_temp["load_case"]=max_temp_comp_case
        
        #%%
        # load cases to be considered
        columns_to_keep=['load_case','Type']
        # (b)
        # filter the  Load_summary_df column to just the load combi of concern
        # max_perm_comp_case
        columns_to_keep.append(max_perm_comp_case)
        Load_summary_df_comp=Load_summary_df.loc[:,columns_to_keep]
        
        # merge the Load_summary_df column with the df_loadcases
        df_loadcases.OutputCase=df_loadcases.OutputCase.astype("str")
        
        # filter out load cases prior to merge
        df_loadcases_b=df_loadcases.loc[(df_loadcases.Joint==max_perm_comp_joint)]
        
        Load_summary_df_comp=Load_summary_df_comp.merge(df_loadcases_b,left_on='load_case',right_on='OutputCase')
        
        Load_summary_df_comp=Load_summary_df_comp[Load_summary_df_comp[max_perm_comp_case]!=0]
        
        Load_summary_df_comp["unfactored-per m run"]=Load_summary_df_comp.F3
        
        Load_summary_df_comp["CASETYPE"]='max_perm_comp'
        Load_summary_df_comp["load_case"]=max_perm_comp_case
        
        #%%
        # load cases to be considered
        columns_to_keep=['load_case','Type']
        # (c)
        # filter the  Load_summary_df column to just the load combi of concern
        # max_tension_comp_
        columns_to_keep.append(max_perm_tension_case)
        Load_summary_df_tension=Load_summary_df.loc[:,columns_to_keep]
        Load_summary_df_tension[max_perm_tension_case]
        
        
        # merge the Load_summary_df column with the df_loadcases
        df_loadcases.OutputCase=df_loadcases.OutputCase.astype("str")
        
        # filter out load cases prior to merge
        df_loadcases_c=df_loadcases.loc[(df_loadcases.Joint==max_perm_tension_joint)]
        
        Load_summary_df_tension=Load_summary_df_tension.merge(df_loadcases_c,left_on='load_case',right_on='OutputCase')
        Load_summary_df_tension=Load_summary_df_tension[Load_summary_df_tension[max_perm_tension_case]!=0]
        
        Load_summary_df_tension["unfactored-per m run"]=Load_summary_df_tension.F3
        
        Load_summary_df_tension["CASETYPE"]='max_tension'
        Load_summary_df_tension["load_case"]=max_perm_tension_case
        
        
        
        #%%
        Load_summary_all=pd.concat([Load_summary_df_temp,Load_summary_df_comp,Load_summary_df_tension],axis=0)
        final_inner=Load_summary_all.groupby(by=["CASETYPE","Type","load_case"]).agg({"unfactored-per m run":sum}).reset_index()
        
        final_inner.load_case.replace({"UPL_ULS":"UPL_SLS"},inplace=True)
        
        # add in UPL_ULS load
        a={"CASETYPE":["max_tension_ULS"],"Type":["DL"],"load_case":[max_perm_tension_case],"unfactored-per m run":[df_loadcombi_inner[df_loadcombi_inner["OutputCase"]=="UPL_ULS"].F3.min()]}
        UPL_ULS=pd.DataFrame(a)
        final_inner=pd.concat([final_inner,UPL_ULS],axis=0)
        
        return final_outer,final_inner
        
        










