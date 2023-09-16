# -*- coding: utf-8 -*-
"""
Created on Thu Jul 20 09:46:53 2023

@author: 10022767
"""


def Creating_box_loads(section_type,loads_cal_path,Joint_coord,connectivity):
    import pandas as pd
    import xlwings as xw
    import numpy as np
    import openpyxl
    
    # for mapping function
    app = xw.App(visible=False)
    wb = app.books.open(loads_cal_path)
    ws = wb.sheets['Wish In Place']
    depth_of_baseSlab_CL= ws.range('U68').value+ws.range('U69').value
    m_bellow_ground=-ws.range('M65').value+ws.range('AC22').value/2
    wb.close()
    app.quit()
    
    
 
    #%% 1) Define functions
    print("1) Define functions")
    # =============================================================================
    # INTERPOLATE function
    def interpolate(i, j, i_coord, j_coord,coord):
        return ((j-i)/(j_coord-i_coord)*(j_coord-coord)-j)
    
    # =============================================================================
    # INTERPOLATE function
    def interpolate_side(top, bot, top_coord, bot_coord,coord):
        return -((bot-top)/(bot_coord-top_coord)*(bot_coord-coord)-bot)

    # 
    # =============================================================================
     
    
    def interpolate_ForLoop_wall(LoadPatterns_set,df,Boundary_col,Boundary_col_value_1st,Boundary_col_value_2nd,joint_loads_case,joint_loads_append,face,i_coord,j_coord):
        
       
        for loadpat in LoadPatterns_set:
            df=df.loc[df["LoadPat"]==str(loadpat)]
            i=df.loc[(df[Boundary_col]==Boundary_col_value_1st)]["load"].values[0]
            j=df.loc[(df[Boundary_col]==Boundary_col_value_2nd)]["load"].values[0] 
            if face=="LHS":
                joint_loads_case["load"]=joint_loads_case["Z"].apply(lambda x: -interpolate(i,j,i_coord,j_coord,x))
        
            elif face=="RHS": 
                joint_loads_case["load"]=joint_loads_case["Z"].apply(lambda x: interpolate(i,j,i_coord,j_coord,x))
        
            joint_loads_case["LoadPat"]=str(loadpat)
            # joint_loads_append=joint_loads_append.append(joint_loads_case,ignore_index=True)
            joint_loads_append=pd.concat([joint_loads_append,joint_loads_case],axis=0)
            
        return joint_loads_append
    
    
    def interpolate_ForLoop_slab(balanced_set,unbalanced_set,df,Boundary_col,joint_loads_case,joint_loads_append,face,i_coord,j_coord):
    
       if face=="roof":
           face="Vertical Roof"
           direction=-1
       elif face=="base":
           face="Vertical Base"
           direction=1
       
       # Balanced loadcases
       
       df=df.loc[df[Boundary_col]==face]
    
       for loadpat in balanced_set:
           if df.loc[(df["LoadPat"]==str(loadpat))].size==0:
               continue
           else:
               i=df.loc[(df["LoadPat"]==str(loadpat))]["load"].values[0]*direction
               j=i
               joint_loads_case["load"]=joint_loads_case["XorR"].apply(lambda x: -interpolate(i,j,i_coord,j_coord,x))
               joint_loads_case["LoadPat"]=str(loadpat)
               # joint_loads_append=joint_loads_append.append(joint_loads_case,ignore_index=True)
               joint_loads_append=pd.concat([joint_loads_append,joint_loads_case],axis=0)
    
       
    
        # unbalanced loadcases
    
       ij_list=[[],[],[]]
       
       for loadpat in unbalanced_set["LHS"]:
           if df.loc[(df["LoadPat"]==str(loadpat))].size==0:
               continue
           else:
        
               i=df.loc[(df["LoadPat"]==str(loadpat))]["load"].values[0]*direction
               ij_list[0].append(i)
          
       for loadpat in unbalanced_set["RHS"]:
           if df.loc[df["LoadPat"]==str(loadpat)].size==0:
               continue
           else:
               j=df.loc[(df["LoadPat"]==str(loadpat))]["load"].values[0]*direction
               ij_list[1].append(j)

               
       for loadpat in unbalanced_set["LHS"]:
           if df.loc[(df["LoadPat"]==str(loadpat))].size==0:
               continue
           else:
               ij_list[2].append(str(loadpat))
     

       for a in range(len((ij_list[0]))):
           joint_loads_case["load"]=joint_loads_case["XorR"].apply(lambda x: -interpolate((ij_list[0][a]),ij_list[1][a],i_coord,j_coord,x))
           joint_loads_case["LoadPat"]=ij_list[2][a]
           # joint_loads_append=joint_loads_append.append(joint_loads_case,ignore_index=True)
           joint_loads_append=pd.concat([joint_loads_append,joint_loads_case])
    
       
       
       return joint_loads_append  
        
    
            
    # =============================================================================
    
    # =============================================================================
    # MAPPING functions
    # translate the baseslab cl to Z=0 to be same as the model 
    def Mapping(df,Boundary_Col,Levels_col):
        index_start=int(df[df[Boundary_Col]=='Box Roof CL'].index.values)
        index_end=int(df[df[Boundary_Col]=='Box Base CL'].index.values)
        df=df.iloc[index_start:index_end+1,:]
        df['Levels_Trans']=(df[Levels_col]+m_bellow_ground)*1000
        
        return df
    
    
    # find closest value
    def find_closest_value(val, df,LoadPat):
        return df.loc[(df['Levels_Trans'] - val).abs().idxmin(),LoadPat]
    
    
    # idnetifying extreme coordinates
    def Identify_end_coords(coords_of_wall_or_slab):
        joint_loads_case=coords_of_wall_or_slab[["Joint","Z"]]
        first_coord=coords_of_wall_or_slab['Z'].min()
        second_coord=coords_of_wall_or_slab['Z'].max()
        
        return first_coord,second_coord,joint_loads_case
    
    
    
    # load cases to rename
    load_case_rename_dict_vert={

        'Load Case 19':'19',
        'Load Case 23':'23',
        'Load Case 39 LHS':'39',
        'Load Case 32':'32',
        'Load Case 36':'36',
        "Load Case 20":'20',
        "Load Case 22":'22',
        "Load Case 39 RHS": "39",
        "Load Case 34": "34",
        "Load Case 53-55":"53",
        "Load Case 4":"4",
        "Load Case 5":"5",
        "Load Case 31":"31",
        "Load Case 31.1":"31.1",
        "Load Case 35":'35',
        "Load Case 37":'37',
        "Load Case 38 LHS":'38',
        "Load Case 41 LHS":'41',
        "Load Case 33":'33',
        "Load Case 40 LHS":'40',
        "LC42 LC43 LHS":'43',
        "LC42 LC43 RHS":'43',
        'Load Case 11':'11', 
        'Load Case 13':'13', 
        'Load Case 15':'15', 
        'Load Case 17':"17",
        'Load Case 17.1':"17.1",
        'Load Case 18 West':'18', 
        'Load Case 21 West':'21',
        'Load Case 56-57':'56'
        }
    
    load_case_rename_dict_hori={
        'Load Case 12':'12',
        'Load Case 14':'14',
        'Load Case 16':'16',
        'Load Case 19':'19',
        'Load Case 23':'23',
        'LC42 LC43 LHS':'42',
        'Load Case 39 LHS':'39',
        'Load Case 32':'32',
        'Load Case 36':'36',
        "Load Case 20":'20',
        "Load Case 22":'22',
        "LC42 LC43 RHS":"42",
        "Load Case 39 RHS": "39",
        "Load Case 34": "34",
        "Load Case 53-55":"53",
        "Load Case 4":"4",
        "Load Case 5":"5",
        "Load Case 35":'35',
        "Load Case 37":'37',
        "Load Case 38 LHS":'38',
        "Load Case 41 LHS":'41',
        "Load Case 33":'33',
        "Load Case 40 LHS":'40',
        'Load Case 11':'11', 
        'Load Case 13':'13', 
        'Load Case 15':'15', 
        'Load Case 18 West':'18', 
        'Load Case 21 West':'21',
        'Load Case 54.1': '54.1',
        'Load Case 54.2': '54.2',
        'Load Case 55.1': '55.1',
        'Load Case 55.2': '55.2',
        'Load Case 56-57':'57'
        
        }
    
    
    
    #%% 2) GENERATING VERTICAL LOADS
    # =======================================================================
    print("2) GENERATING VERTICAL LOADS")
    
    
    Joint_coord['Z']=pd.to_numeric(Joint_coord['Z'])
    
    

    # LHS earth pressure table
    LHS_df=pd.read_excel(loads_cal_path,usecols="AB:BJ",header=54)
    LHS_df=LHS_df.iloc[:24,:]
    # RHS earth pressure table
    RHS_df=pd.read_excel(loads_cal_path,usecols="AB:BA",header=81)
    RHS_df=RHS_df.iloc[:21,:]
    # water table
    water_df=pd.read_excel(loads_cal_path,usecols="AC:BH",header=123)
    water_df=water_df.rename(columns=lambda x: x.strip())
    water_df=water_df.iloc[:24,:]
        
    
    # Locating coordinates from SAP2000 model
    # filtering the left wall
    Joint_coord_left=Joint_coord.loc[Joint_coord["XorR"]==Joint_coord["XorR"].min()]
    
    # filtering the right wall
    Joint_coord_right=Joint_coord.loc[Joint_coord["XorR"]==Joint_coord["XorR"].max()]
    
    # filtering the roof
    Joint_coord_roof=Joint_coord.loc[Joint_coord["Z"]==Joint_coord["Z"].max()]
    
    # filtering the base slab
    Joint_coord_base=Joint_coord.loc[Joint_coord["Z"]==Joint_coord["Z"].min()]
    

    balanced_set=(
        "Load Case 31",
        "Load Case 31.1",
        "Load Case 33",
        "Load Case 34",
        "Load Case 35",
        "Load Case 37",
        )
    
    
    
    unbalanced_set={
        "LHS":(
            "Load Case 38 LHS",
            "Load Case 40 LHS",
            "Load Case 41 LHS",
            "LC42 LC43 LHS"
            ),
        "RHS":(
            "Load Case 38 RHS",
            "Load Case 40 RHS",
            "Load Case 41 RHS",
            "LC42 LC43 RHS"
            )
        }
    
    
    
    # ______________________
    # vertical water(roof,base)
    # roof slab and base slab
    water_df_roof=water_df[[
        "Boundary",
        "Load Case 31",
        "Load Case 31.1",
        "Load Case 34",
        "Load Case 35",
        "Load Case 37",
        "Load Case 38 LHS",
        "Load Case 41 LHS",
        "Load Case 38 RHS",
        "Load Case 41 RHS"
        ]]
    
    
    
    water_df_base=water_df[[
        "Boundary",
        "Load Case 33",
        "Load Case 34",
        "Load Case 35",
        "Load Case 37",
        "Load Case 40 LHS",
        "LC42 LC43 LHS",
        "Load Case 40 RHS",
        "LC42 LC43 RHS",
        ]]
    
    
    # unpivoting 
    water_df_roof_unpivot=water_df_roof.melt(id_vars=['Boundary'], var_name='LoadPat', value_name='load')
    water_df_base_unpivot=water_df_base.melt(id_vars=['Boundary'], var_name='LoadPat', value_name='load')
    
    # Roof Coordinate
    # Applying joint loads vertical water(roof,base) on roof slab
    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT
    
    df=water_df_roof_unpivot
    Boundary_col="Boundary"
    
    
    # DETERMINE THE LOAD AT THE ROOF COORDINATE JOINT
    joint_loads_append_v=pd.DataFrame() #creating an empty dataframe
    
    coords_of_wall_or_slab=Joint_coord_roof
    joint_loads_case=coords_of_wall_or_slab[["Joint","XorR"]]
    first_coord=coords_of_wall_or_slab['XorR'].min()
    second_coord=coords_of_wall_or_slab['XorR'].max()
    
    
    face="roof"
    joint_loads_append_v=interpolate_ForLoop_slab(balanced_set,unbalanced_set,df,Boundary_col,joint_loads_case,joint_loads_append_v,face,first_coord,second_coord)
    
    # Base Coordinate
    # Applying joint loads vertical water(roof,base) on base slab
    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT
    
    df=water_df_base_unpivot
    Boundary_col="Boundary"
    
    # DETERMINE THE LOAD AT THE ROOF COORDINATE JOINT
    coords_of_wall_or_slab=Joint_coord_base
    joint_loads_case=coords_of_wall_or_slab[["Joint","XorR"]]
    first_coord=coords_of_wall_or_slab['XorR'].min()
    second_coord=coords_of_wall_or_slab['XorR'].max()
    
    face="base"
    joint_loads_append_v=interpolate_ForLoop_slab(balanced_set,unbalanced_set,df,Boundary_col,joint_loads_case,joint_loads_append_v,face,first_coord,second_coord)
    

    # Vertical Soil(roof)
    
    balanced_set=(
        "Load Case 11",
        "Load Case 13",
        "Load Case 15",
        "Load Case 17",
        "Load Case 17.1"
        )
    
    unbalanced_set={
        "LHS":(
            "Load Case 18 West",
            "Load Case 21 West"
            ),
        "RHS":(
            "Load Case 18 East",
            "Load Case 21 East"
            )
        }
    
    # roof slab
    vert_soil_df_left=LHS_df[[
        "Boundary LHS",
        "Load Case 11",
        "Load Case 13",
        "Load Case 15",
        "Load Case 17",
        "Load Case 17.1",
        "Load Case 18 West",
        "Load Case 21 West"
        ]]
    
    vert_soil_df_right=RHS_df[[
        "Boundary RHS",
        "Load Case 18 East",
        "Load Case 21 East"
        ]]
    
    vert_soil_df=vert_soil_df_left.merge(vert_soil_df_right,left_on="Boundary LHS",right_on="Boundary RHS").drop(axis=1,columns="Boundary RHS").rename(columns={"Boundary LHS":"Boundary"})
    
    # unpivoting 
    vert_soil_df_unpivot=vert_soil_df.melt(id_vars=['Boundary'], var_name='LoadPat', value_name='load')
    
    
    # Roof Coordinate
    # Applying joint loads vertical water(roof,base) on roof slab
    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT
    
    # =============================================================================
    # df=water_df_roof_unpivot
    # Boundary_col="Boundary"
    # =============================================================================
    
    
    # DETERMINE THE LOAD AT THE ROOF COORDINATE JOINT
    coords_of_wall_or_slab=Joint_coord_roof
    joint_loads_case=coords_of_wall_or_slab[["Joint","XorR"]]
    first_coord=coords_of_wall_or_slab['XorR'].min()
    second_coord=coords_of_wall_or_slab['XorR'].max()
    
    
    
    # Balanced loadcases
    direction=-1
    vert_soil_df_unpivot=vert_soil_df_unpivot.loc[vert_soil_df_unpivot["Boundary"]=="Vertical Roof"]
    for loadpat in balanced_set:
        if (vert_soil_df_unpivot["LoadPat"]==str(loadpat)).size==0:
            continue
        else:
            i=vert_soil_df_unpivot.loc[(vert_soil_df_unpivot["LoadPat"]==str(loadpat))]["load"].values[0]*direction
            j=i
            joint_loads_case["load"]=joint_loads_case["XorR"].apply(lambda x: -interpolate(i,j,first_coord,second_coord,x))
            joint_loads_case["LoadPat"]=str(loadpat)
            # joint_loads_append_v=joint_loads_append_v.append(joint_loads_case,ignore_index=True)
            joint_loads_append_v=pd.concat([joint_loads_append_v,joint_loads_case],axis=0)
    
    
    
     # unbalanced loadcases
    
    
    ij_list=[[],[],[]]
    for loadpat in unbalanced_set["LHS"]:
        if (vert_soil_df_unpivot["LoadPat"]==str(loadpat)).size==0:
            continue
        else:
     
            i=vert_soil_df_unpivot.loc[(vert_soil_df_unpivot["LoadPat"]==str(loadpat))]["load"].values[0]*direction
            ij_list[0].append(i)
       
    for loadpat in unbalanced_set["RHS"]:
        if (vert_soil_df_unpivot["LoadPat"]==str(loadpat)).size==0:
            continue
        else:
            j=vert_soil_df_unpivot.loc[(vert_soil_df_unpivot["LoadPat"]==str(loadpat))]["load"].values[0]*direction
            ij_list[1].append(j)
            
    for loadpat in unbalanced_set["LHS"]:
        if (vert_soil_df_unpivot["LoadPat"]==str(loadpat)).size==0:
            continue
        else:
            ij_list[2].append(str(loadpat))
    
    for a in range(len((ij_list[0]))):
        joint_loads_case["load"]=joint_loads_case["XorR"].apply(lambda x: -interpolate((ij_list[0][a]),ij_list[1][a],first_coord,second_coord,x))
        joint_loads_case["LoadPat"]=ij_list[2][a]
        # joint_loads_append_v=joint_loads_append_v.append(joint_loads_case,ignore_index=True)
        joint_loads_append_v=pd.concat([joint_loads_append_v,joint_loads_case],axis=0)

    # Vertical Surcharge(roof)
    wb=openpyxl.load_workbook(loads_cal_path)
    sheet=wb.get_sheet_by_name("Wish In Place")

    

# for roof slab
    dict_bal_V_LC_r={
     '2': -sheet['R81'].value,  # SDL roof
     '52': -sheet['AD44'].value,  # aircraft load
     '56': -sheet['AD46'].value, # construction load
     '3': -sheet['AD47'].value,  # Pavement selfweight
     }
    
# for base  slab
    dict_bal_V_LC_b={
     '2': -sheet['R82'].value,  # SDL baseslab
     '51': -sheet['AD45'].value  # Internal Live load
     }
    

    
    # ROOF slab loading
    # DETERMINE THE LOAD AT THE ROOF COORDINATE JOINT
    # locate joints
    coords_of_wall_or_slab=Joint_coord_roof
    joint_loads_case=coords_of_wall_or_slab[["Joint","XorR"]]
    first_coord=coords_of_wall_or_slab['XorR'].min()
    second_coord=coords_of_wall_or_slab['XorR'].max()
    
    for loadcase, load in dict_bal_V_LC_r.items():
        i=load
        j=i #left and right same
        joint_loads_case["load"]=joint_loads_case["XorR"].apply(lambda x: -interpolate(i,j,first_coord,second_coord,x))
        joint_loads_case["LoadPat"]=loadcase
        # joint_loads_append_v=joint_loads_append_v.append(joint_loads_case,ignore_index=True)
        joint_loads_append_v=pd.concat([joint_loads_append_v,joint_loads_case],axis=0)
        
        
    # Base slab loading
    # DETERMINE THE LOAD AT THE BASE COORDINATE JOINT
    # locate joints
    coords_of_wall_or_slab=Joint_coord_base
    joint_loads_case=coords_of_wall_or_slab[["Joint","XorR"]]
    first_coord=coords_of_wall_or_slab['XorR'].min()
    second_coord=coords_of_wall_or_slab['XorR'].max()
    
    for loadcase, load in dict_bal_V_LC_b.items():
        i=load
        j=i #left and right same
        joint_loads_case["load"]=joint_loads_case["XorR"].apply(lambda x: -interpolate(i,j,first_coord,second_coord,x))
        joint_loads_case["LoadPat"]=loadcase
        # joint_loads_append_v=joint_loads_append_v.append(joint_loads_case,ignore_index=True)
        joint_loads_append_v=pd.concat([joint_loads_append_v,joint_loads_case],axis=0)
    
    #%% 3) GENERATING HORIZONTAL LOADS
    # =============================================================================
    print("3) GENERATING HORIZONTAL LOADS")
    
    # =============================================================================
    # #Interpolation
    # 1)getting loads from extreme ends
    # 2)map these extreme loads to the model node position and interpolating them to get the individual node load for their repective position
    # =============================================================================
    
    water_df=pd.read_excel(loads_cal_path,usecols="AC:BF",header=123)
    water_df=water_df.rename(columns=lambda x: x.strip())
    water_df=water_df.iloc[:13,:]
    
    # water pressure LHS
    water_df_LHS=water_df[[
        "Boundary",
        "Load Case 32",
        "Load Case 36",
        "Load Case 34",
        "LC42 LC43 LHS",
        "Load Case 39 LHS",
        ]]
    
    # water pressure RHS
    water_df_RHS=water_df[[
        "Boundary",
        "Load Case 32",
        "Load Case 36",
        "Load Case 34",
        "LC42 LC43 RHS",
        "Load Case 39 RHS"
        ]]
    
    water_df_LHS=water_df_LHS.rename(columns=lambda x: x.strip())
    water_df_LHS=water_df_LHS.rename(columns=load_case_rename_dict_hori)
    water_df_LHS=water_df_LHS[(water_df_LHS["Boundary"]=="Box Roof CL") | (water_df_LHS["Boundary"]=="Box Base CL")]
    water_df_LHS=water_df_LHS.dropna(axis='columns',how='all')
    water_df_LHS_unpivot=water_df_LHS.melt(id_vars=['Boundary'], var_name='LoadPat', value_name='load')
    
    
    water_df_RHS=water_df_RHS.rename(columns=lambda x: x.strip())
    water_df_RHS=water_df_RHS.rename(columns=load_case_rename_dict_hori)
    water_df_RHS=water_df_RHS[(water_df_RHS["Boundary"]=="Box Roof CL") | (water_df_RHS["Boundary"]=="Box Base CL")]
    water_df_RHS=water_df_RHS.dropna(axis='columns',how='all')
    water_df_RHS_unpivot=water_df_RHS.melt(id_vars=['Boundary'], var_name='LoadPat', value_name='load')

    
    
    
    # =============================================================================
    # # LHS
    # =============================================================================
    
    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT
    end_coords_results=Identify_end_coords(coords_of_wall_or_slab=Joint_coord_left)
    joint_loads_case=end_coords_results[2]
    # prepare empty df for joint_loads
    joint_loads_append_h=pd.DataFrame()
    
    # identify top and bottom joint coordinate
    bot_coord=end_coords_results[0]
    top_coord=end_coords_results[1]
    
    LHS_chosen_loadpat_water=['32', '36', '34', '42','39']
    for LoadPat in LHS_chosen_loadpat_water:
        top=water_df_LHS_unpivot.loc[(water_df_LHS_unpivot["Boundary"]=="Box Roof CL") & (water_df_LHS_unpivot["LoadPat"]==LoadPat)]["load"].values[0]
        bot=water_df_LHS_unpivot.loc[(water_df_LHS_unpivot["Boundary"]=="Box Base CL") & (water_df_LHS_unpivot["LoadPat"]==LoadPat)]["load"].values[0] 
        joint_loads_case["load"]=joint_loads_case["Z"].apply(lambda x: interpolate_side(top,bot,top_coord,bot_coord,x))
        joint_loads_case["LoadPat"]=LoadPat
        # joint_loads_append_h=joint_loads_append_h.append(joint_loads_case,ignore_index=True)
        joint_loads_append_h=pd.concat([joint_loads_append_h,joint_loads_case],axis=0)
        
    # =============================================================================
    # # RHS
    # =============================================================================
    
    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT
    end_coords_results=Identify_end_coords(coords_of_wall_or_slab=Joint_coord_right)
    joint_loads_case=end_coords_results[2]
    # identify top and bottom joint coordinate
    bot_coord=end_coords_results[0]
    top_coord=end_coords_results[1]
    
    RHS_chosen_loadpat_water=['32', '36', '34', '42','39']
    for LoadPat in RHS_chosen_loadpat_water:
        top=water_df_RHS_unpivot.loc[(water_df_RHS_unpivot["Boundary"]=="Box Roof CL") & (water_df_RHS_unpivot["LoadPat"]==LoadPat)]["load"].values[0]
        bot=water_df_RHS_unpivot.loc[(water_df_RHS_unpivot["Boundary"]=="Box Base CL") & (water_df_RHS_unpivot["LoadPat"]==LoadPat)]["load"].values[0] 
        joint_loads_case["load"]=-joint_loads_case["Z"].apply(lambda x: interpolate_side(top,bot,top_coord,bot_coord,x)) #right have -ve factor for opposite direction
        joint_loads_case["LoadPat"]=LoadPat
        # joint_loads_append_h=joint_loads_append_h.append(joint_loads_case,ignore_index=True)
        joint_loads_append_h=pd.concat([joint_loads_append_h,joint_loads_case],axis=0)
        
    # =============================================================================
    # # Matching closest point
    # Matching closest point
    # 1)find abosulte closest position of model node and load calculation table and matching them 
    
    # # ________________________
    # # lateral soil(LHS,RHS) including surcharge

    LHS_df=pd.read_excel(loads_cal_path,usecols="AB:BI",header=54)
    LHS_df=LHS_df.iloc[:22,:]
    LHS_df=LHS_df[[
        "Boundary LHS",
        "Level LHS",
        "Thickness to layer above",
        "Load Case 12",
        "Load Case 14",
        "Load Case 16",
        "Load Case 19",
        "Load Case 23",
        "Load Case 4",
        "Load Case 5",
        'Load Case 54.1',
        'Load Case 54.2',
        "Load Case 53-55", #for 53 lateral surcharge, 54 and 55 taken care of
        "Load Case 56-57", #only to take 57 for lateral
        ]]
    
    
    
    RHS_df=pd.read_excel(loads_cal_path,usecols="AB:BG",header=81)
    RHS_df=RHS_df.iloc[:19,:]
    RHS_df=RHS_df[[
        "Boundary RHS",
        "Level RHS",
        "Thickness to layer above",
        "Load Case 12",
        "Load Case 14",
        "Load Case 16",
        "Load Case 20",
        "Load Case 22",
        "Load Case 4",
        "Load Case 5",
        'Load Case 55.1',
        'Load Case 55.2',
        "Load Case 53-55", #for 53 lateral surcharge, 54 and 55 taken care of
        "Load Case 56-57", #only to take 57 for lateral
        
        ]]
    
    # LHS
    LHS_df=LHS_df.rename(columns=lambda x: x.strip())
    LHS_df=LHS_df.rename(columns=load_case_rename_dict_hori)
    LHS_df=Mapping(df=LHS_df,Boundary_Col='Boundary LHS',Levels_col='Level LHS')    
    LHS_df.iloc[:,1:]=LHS_df.iloc[:,1:].apply(pd.to_numeric, errors='coerce')
    # RHS
    RHS_df=RHS_df.rename(columns=lambda x: x.strip())
    RHS_df=RHS_df.rename(columns=load_case_rename_dict_hori)
    RHS_df=Mapping(df=RHS_df,Boundary_Col='Boundary RHS',Levels_col='Level RHS')
    RHS_df.iloc[:,1:]=RHS_df.iloc[:,1:].apply(pd.to_numeric, errors='coerce')



    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT for LHS
    end_coords_results=Identify_end_coords(coords_of_wall_or_slab=Joint_coord_left)
    joint_loads_case=end_coords_results[2]
    # selecting LoadPat and applying loads to for each lc
    LHS_chosen_loadpat=['12', '14', '16', '19','23', '4', '5', '54.1', '54.2', '53','57']
    for LoadPat in LHS_chosen_loadpat:
        # Apply the custom function to each value in column2 and store the result in a new column
        joint_loads_case['load'] = joint_loads_case['Z'].apply(lambda x: find_closest_value(x, LHS_df,LoadPat))
        joint_loads_case["LoadPat"]=LoadPat
        # joint_loads_append_h=joint_loads_append_h.append(joint_loads_case,ignore_index=True)
        joint_loads_append_h=pd.concat([joint_loads_append_h,joint_loads_case],axis=0)
        
    # DETERMINE THE LOAD AT THE Z COORDINATE JOINT for RHS
    end_coords_results=Identify_end_coords(coords_of_wall_or_slab=Joint_coord_right)
    joint_loads_case=end_coords_results[2]    
    # selecting LoadPat and applying loads to for each lc
    RHS_chosen_loadpat=['12', '14', '16', '20','22', '4', '5', '55.1', '55.2', '53', '57']
    for LoadPat in RHS_chosen_loadpat:
        # Apply the custom function to each value in column2 and store the result in a new column
        joint_loads_case['load'] = -joint_loads_case['Z'].apply(lambda x: find_closest_value(x, RHS_df,LoadPat))
        joint_loads_case["LoadPat"]=LoadPat
        # joint_loads_append_h=joint_loads_append_h.append(joint_loads_case,ignore_index=True)
        joint_loads_append_h=pd.concat([joint_loads_append_h,joint_loads_case],axis=0)


    
    # defining horizontal loads joints
    joint_loads_append_h=joint_loads_append_h.astype({"Joint":'str'})

    #%% 4) COMBINING VERTICAL AND HORIZONTAL LOADS AND APPLYING TO FRAME ELEMENT IN MODEL
    # =============================================================================
    print("4) COMBINING VERTICAL AND HORIZONTAL LOADS AND APPLYING TO FRAME ELEMENT IN MODEL")
    # rename 
    joint_loads_append_h['LoadPat'] = joint_loads_append_h['LoadPat'].replace(load_case_rename_dict_hori)
    joint_loads_append_v['LoadPat'] = joint_loads_append_v['LoadPat'].replace(load_case_rename_dict_vert)
    # ASSIGN LOAD TO THE JOINT(Connectivity - Frame) and prepare final loads table 
    
    # MAP THE JOINTS TO THE FRAME(Connectivity - Frame)
    # UNPIVOT THE JOINTI AND JOINTJ (Connectivity - Frame)
    connectivity_unpivot=connectivity.iloc[:,:3].melt(id_vars=['Frame'], var_name='Joint_Ref', value_name='Joint')
    connectivity_unpivot["Joint_Ref"]=connectivity_unpivot["Joint_Ref"].replace("JointI","RelDistA")
    connectivity_unpivot["Joint_Ref"]=connectivity_unpivot["Joint_Ref"].replace("JointJ","RelDistB")
    connectivity_unpivot=connectivity_unpivot.astype({"Joint":'str'})
    
    
    # Merge the loads vertical
    Loads_combined_v=connectivity_unpivot.merge(joint_loads_append_v[["Joint","load","LoadPat"]],on="Joint")
    Loads_combined_v.drop("Joint",axis=1,inplace=True)
    
    
    # PIVOT BACK THE JOINT (Frame Loads - Distributed)
    final_v=pd.pivot_table(Loads_combined_v,values="load",columns=["Joint_Ref"],index=["Frame","LoadPat"]).reset_index()
    
    final_v=final_v.astype({"Frame":'str'})
    final_v.rename(columns={"RelDistA": "FOverLA", "RelDistB": "FOverLB"},inplace=True)
    

    
    # Frame_loads_v=Frame_loads.append(final_v,ignore_index=True)
    Frame_loads_v=final_v
    Frame_loads_v["Dir"]="Z"

    
    # Merge the loads horizontal
    Loads_combined_h=connectivity_unpivot.merge(joint_loads_append_h[["Joint","load","LoadPat"]],on="Joint")
    Loads_combined_h.drop("Joint",axis=1,inplace=True)
    
    
    # PIVOT BACK THE JOINT (Frame Loads - Distributed)
    final_h=pd.pivot_table(Loads_combined_h,values="load",columns=["Joint_Ref"],index=["Frame","LoadPat"]).reset_index()
    
    final_h=final_h.astype({"Frame":'str'})
    final_h.rename(columns={"RelDistA": "FOverLA", "RelDistB": "FOverLB"},inplace=True)
    
    # Frame_loads_h=Frame_loads.append(final_h,ignore_index=True)
    Frame_loads_h=final_h
    Frame_loads_h["Dir"]="X"

    
    # Combined frame loads
    # Frame_loads=Frame_loads_v.append(Frame_loads_h,ignore_index=True)
    Frame_loads=pd.concat([Frame_loads_v,Frame_loads_h],axis=0)
    
    Frame_loads=Frame_loads.astype({"Frame":'str'})
    
    Frame_loads[['FOverLA','FOverLB']]=Frame_loads[['FOverLA','FOverLB']].round(1)
    Frame_loads[['FOverLA','FOverLB']]=Frame_loads[['FOverLA','FOverLB']].fillna(0)
    
   
    # # extract LC2 and 51 and append
    # Frame_loads_defaultset=pd.read_excel(loads_cal_path,sheet_name="LC2 and 51",header=1, dtype={'LoadPat': str,"Frame": str}).iloc[1:]
    # Frame_loads_defaultset[["AbsDistA","AbsDistB"]]=np.nan
    Frame_loads.FOverLA=Frame_loads.FOverLA.apply(lambda x: float(x))
    Frame_loads.FOverLB=Frame_loads.FOverLB.apply(lambda x: float(x))
    
    # Frame_loads=Frame_loads.merge(Frame_loads_defaultset[["Frame","LoadPat","FOverLA","FOverLB",'Dir']],on=["Frame","LoadPat"],how="outer")
    # Frame_loads['Dir_x']=Frame_loads['Dir_x'].fillna("")
    # Frame_loads['Dir_y']=Frame_loads['Dir_y'].fillna("")
    # Frame_loads['Dir']=(Frame_loads['Dir_x']+Frame_loads['Dir_y']).str[0]
    
    # Frame_loads['FOverLA']=Frame_loads.FOverLA_y.fillna(0)+Frame_loads.FOverLA_x.fillna(0)
    # Frame_loads['FOverLB']=Frame_loads.FOverLB_y.fillna(0)+Frame_loads.FOverLB_x.fillna(0)
   
    Frame_loads["CoordSys"]="GLOBAL"
    Frame_loads["Type"]="Force"
    Frame_loads["DistType"]='RelDist'
    Frame_loads["RelDistA"]=0
    Frame_loads["RelDistB"]=1
    Frame_loads["AbsDistA"]=np.nan
    Frame_loads["AbsDistB"]=np.nan
    
    Frame_loads.reset_index(inplace=True,drop=True)
    
    def convert(Dir):
        if Dir=='Z' or Dir=='z': # see SAP2000 source code for direction reference
            return 6
        elif Dir=='X' or Dir=='x':  # see SAP2000 source code for direction reference
            return 4
        elif Dir=='Y' or Dir=='y':   # see SAP2000 source code for direction reference
            return 5

    Frame_loads["Dir"]
    Frame_loads.Dir=Frame_loads.Dir.apply(lambda x: convert(x))

    Frame_loads=Frame_loads[['Frame', 'LoadPat', 'Dir', 'FOverLA', 'FOverLB', 'CoordSys', 'Type', 'DistType',
           'RelDistA', 'RelDistB', 'AbsDistA', 'AbsDistB']]
    
    
    
    return Frame_loads





