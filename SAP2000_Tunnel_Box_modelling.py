# -*- coding: utf-8 -*-
"""
Created on Sat Aug 26 23:55:46 2023

@author: 10022767
"""

import os
import win32com.client
import sys
import pandas as pd
import comtypes.client
import openpyxl
import xlwings as xw
from Extracting_Sap2000_model_info import get_tunnel_sapmodel_info
import numpy as np
import Getting_box_loads5_Ver3
import math

section_type="Test"



Parent_folder=r"C:\Users\10022767\Surbana Jurong Private Limited(1)\Johanna Enriquez - C-SG-003399_CAG Intra Tunnel\Calculations\CST"
# path for tunnel excel calculation
loads_cal_path=Parent_folder+"\Section "+section_type+"\Load Calculation and Spring-Section "+section_type+".xlsx"
load_combi_summary_path=Parent_folder+"\Model information\Load combinatios CST.xlsx"
print("Model Initializing")
model_path=Parent_folder+"\Section "+section_type+"\SAP2000 Model\\"+section_type+".sdb"
SapObject = win32com.client.Dispatch("CSI.SAP2000.API.SapObject")
SapObject.ApplicationStart()
SapModel = SapObject.SapModel  # create SAP2000 model object
SapModel.InitializeNewModel() #initialize model



#%% *CREATING NEW MODEL
#create new blank model
ret = SapModel.File.NewBlank()
SapModel.SetPresentUnits(6)
#define material property

MATERIAL_CONCRETE = 2
ret = SapModel.PropMaterial.SetMaterial('CONC', MATERIAL_CONCRETE)


# sets the model global degrees of freedom to plane frame
SapModel.Analyze.SetActiveDOF([True, False, True, False, True, False])

# loadcases/combi
Load_summary_df=pd.read_excel(load_combi_summary_path,sheet_name="Load summary(updated)",header=1,converters={'load_case': str})
Load_cases=list(Load_summary_df.load_case)
Load_combi=list(Load_summary_df.columns[2:24])

Load_summary_df['Type2']=Load_summary_df['Type'].apply(lambda x: 1 if x=="DL" else 3)
tuple_of_loadcase= tuple(Load_summary_df.apply(lambda row: (row['load_case'], row['Type2']), axis=1))

#table summary of load_Scenarios to consider
for t_loadcase in tuple_of_loadcase:
    if t_loadcase[0]=="1":
        SelfWTMultiplier=1
        MyType=t_loadcase[1]
    else:
        SelfWTMultiplier=0
        MyType=t_loadcase[1]
        
    print(t_loadcase[0])    
    SapModel.LoadPatterns.Add(Name=t_loadcase[0],MyType=MyType,SelfWTMultiplier=SelfWTMultiplier)
    SapModel.LoadCases.StaticNonlinear.SetCase(Name=t_loadcase[0])
    SapModel.LoadCases.StaticNonlinear.SetLoads(Name=t_loadcase[0],NumberLoads=1,LoadName=(t_loadcase[0],),LoadType= ("Load",),SF = (1,))


tuple_of_LoadName=tuple(Load_summary_df.loc[:,'load_case'].to_list())
for load_combi in Load_summary_df.loc[:,"101":"UPL_SLS"].columns:
    tuple_of_LoadType=tuple(len(tuple_of_LoadName)*["Load"])
    tuple_of_SF=tuple(Load_summary_df.loc[:,load_combi].to_list())
    print(SapModel.LoadCases.StaticNonlinear.SetCase(Name=load_combi+"-NL")) #create nonlinear load cases combi
    print(SapModel.LoadCases.StaticNonlinear.SetLoads(Name=load_combi+"-NL",NumberLoads=len(tuple_of_LoadName),LoadName=tuple_of_LoadName,LoadType= tuple_of_LoadType,SF = tuple_of_SF))   #for adding non linear load combi

    print(SapModel.RespCombo.Add(load_combi,0))
    print(SapModel.RespCombo.SetCaseList(load_combi,0,load_combi+"-NL",1))# for adding loadcombi
    
# ModelPath=r"C:\Users\10022767\OneDrive\Desktop\Tunnel_testing.sdb"
# SapModel.File.Save(ModelPath)



#%% *Drawing the model
app = xw.App(visible=False)
wb = app.books.open(loads_cal_path)
ws = wb.sheets['Wish In Place']

External_Wall_thickness= ws.range('AC16').value
Internal_Wall_thickness= ws.range('AC17').value
first_cell_width= ws.range('AC18').value
second_cell_width= ws.range('AC19').value 
third_cell_width= ws.range('AC20').value
roof_slab_thickness= ws.range('AC21').value
depth_of_tunnel= ws.range('U68').value #depth from ground level to centre of roof slab thickness
depth_of_baseSlab_CL= ws.range('U68').value+ws.range('U69').value # for mapping function
base_slab_thickness= ws.range('AC22').value
baseSlab_CL_SHD= ws.range('M65').value-base_slab_thickness /2
internal_height= ws.range('AC23').value
# Int_pile_position=ws.range('AC24').value
Int_pile_position=ws.range('AC24').value
no_of_cells=ws.range('AC15').value
length_element=ws.range('AC25').value
Wind_pressure_internal=ws.range('AD48').value

wb.close()
app.quit()


#assign isotropic mechanical properties to material
ret = SapModel.PropMaterial.SetMPIsotropic('CONC', 3600, 0.2, 0.0000055)
#define rectangular frame section property
print(SapModel.PropFrame.SetRectangle('R1', 'CONC', 1, 1)) #external wall, base slab and roof
print(SapModel.PropFrame.SetRectangle('R2', 'CONC', 0.6, 1)) #internal wall
ret = SapModel.PropFrame.SetRectangle('NULL BEAM', 'CONC', 0.1, 0.1)



#define frame section property modifiers
ModValue = [1, 1, 1, 1, 1, 1, 1, 1]
ret = SapModel.PropFrame.SetModifiers('R1', ModValue)

Slab_BaseRoof_dict={}

#add frame object by coordinates

x1=0
if second_cell_width>0:
    x2=x1+first_cell_width+Internal_Wall_thickness/2
    x3=x2+Internal_Wall_thickness/2+second_cell_width+Internal_Wall_thickness/2
    
else:
    x2=x1+first_cell_width
    
if third_cell_width>0:
    x4=x3+Internal_Wall_thickness/2+third_cell_width+External_Wall_thickness/2
else:
    pass

z1=0
z2=0+internal_height+roof_slab_thickness/2

# Base slab
[ret,BS1] = SapModel.FrameObj.AddByCoord( 0 , 0, 0 , x2, 0, 0, '', 'R1', '', 'Global')
Slab_BaseRoof_dict[BS1]=first_cell_width
if second_cell_width>0:
    [ret,BS2] = SapModel.FrameObj.AddByCoord( x2 , 0, 0 , x3, 0, 0, '', 'R1', '', 'Global')
    Slab_BaseRoof_dict[BS2]=second_cell_width
else:
    pass
if third_cell_width>0:
    [ret,BS3] = SapModel.FrameObj.AddByCoord( x3, 0, 0 , x4, 0, 0, '', 'R1', '', 'Global')
    Slab_BaseRoof_dict[BS3]=third_cell_width
else:
    pass

# Roof slab
[ret,RS1] = SapModel.FrameObj.AddByCoord( 0 , 0, z2 , x2, 0, z2, '', 'R1', '', 'Global')
Slab_BaseRoof_dict[RS1]=first_cell_width
if second_cell_width!=0:
    [ret,RS2]= SapModel.FrameObj.AddByCoord( x2 , 0, z2 , x3, 0, z2, '', 'R1', '', 'Global')
    Slab_BaseRoof_dict[RS2]=second_cell_width
else:
    pass
if third_cell_width!=0:
    [ret,RS3] = SapModel.FrameObj.AddByCoord( x3, 0, z2 , x4, 0, z2, '', 'R1', '', 'Global')
    Slab_BaseRoof_dict[RS3]=third_cell_width
else:
    pass

Walls_list=[]
# walls
[ret,W1]= SapModel.FrameObj.AddByCoord( x1 ,0 ,0, x1, 0, z2, '', 'R1', '', 'Global')
Walls_list.append(W1)
[ret,W2] = SapModel.FrameObj.AddByCoord( x2 , 0, 0 ,x2, 0, z2, '', 'R2', '', 'Global')
Walls_list.append(W2)
if second_cell_width!=0:
    [ret,W3] = SapModel.FrameObj.AddByCoord( x3 , 0, 0 , x3, 0, z2, '', 'R2', '', 'Global')
    Walls_list.append(W3)
else:
    pass
if third_cell_width!=0:
    [ret,W4] = SapModel.FrameObj.AddByCoord( x4 , 0, 0 , x4, 0, z2, '', 'R1', '', 'Global')
    Walls_list.append(W4)

# # subdivide frames
for frame_element,width in Slab_BaseRoof_dict.items():
    SapModel.EditFrame.DivideByRatio(Name=frame_element,Num=math.ceil(width/length_element),Ratio=1)

for frame_element in Walls_list:
    SapModel.EditFrame.DivideByRatio(Name=frame_element,Num=math.ceil(internal_height/length_element),Ratio=1)    
    

# idenitify all joints and frames
All_joints=SapModel.PointObj.GetNameList()[2]
All_frames=SapModel.FrameObj.GetNameList()[2]

# assign pile springs
 # for "Joint Coordinates"
list_of_joint_coords=[]
for joint in All_joints:
    a=list(SapModel.PointObj.GetCoordCartesian(Name=joint)[1:])
    a.insert(0,joint)
    list_of_joint_coords.append(a)
df_joint_coords=pd.DataFrame(list_of_joint_coords,columns=["Joint","XorR",'Y',"Z"])
df_joint_coords[["XorR",'Y',"Z"]]=df_joint_coords[["XorR",'Y',"Z"]].round(4)

base_coord=df_joint_coords.Z.min()
left_coord=df_joint_coords.XorR.min()
right_coord=df_joint_coords.XorR.max()

ext_pile_spring_value=[0,0,265000,0,0,0]
int_pile_spring_value=[0,0,265000,0,0,0]

ext_pile_joint=df_joint_coords.loc[(df_joint_coords.Z==base_coord) &((df_joint_coords.XorR==left_coord)|(df_joint_coords.XorR==right_coord)) ].Joint
SapModel.PointObj.SetSpring(Name=ext_pile_joint.iloc[0],K=ext_pile_spring_value,Replace=True)
SapModel.PointObj.SetSpring(Name=ext_pile_joint.iloc[1],K=ext_pile_spring_value,Replace=True)


if isinstance(Int_pile_position, (float,int)):
    x_closest=df_joint_coords.loc[(df_joint_coords['XorR'] - Int_pile_position).abs().idxmin(),'XorR']
    int_pile_joint=df_joint_coords.loc[(df_joint_coords.Z==base_coord) &(df_joint_coords.XorR==x_closest)].Joint.iloc[0]
    SapModel.PointObj.SetSpring(Name=int_pile_joint,K=int_pile_spring_value,Replace=True)

# save model
SapModel.File.Save(model_path)
print("Model Saved")





#%% EXTRACTING MODEL INFORMATION

model_info_result=get_tunnel_sapmodel_info(SapModel)
df_connectivity=model_info_result[0]
df_frames=model_info_result[1]
df_joint_coords=model_info_result[2]
df_spring_joints=model_info_result[3]
df_joint_coords[["XorR",'Y',"Z"]]=df_joint_coords[["XorR",'Y',"Z"]].round(4)
# identify pile joints
pile_joints=list(df_spring_joints[df_spring_joints.U3>0].Joint)

# filter out pile joints
# identify the outerpiles(min and max X coord)
df_joint_coords_piles=df_joint_coords[df_joint_coords["Joint"].isin(pile_joints)]
df_joint_coords_piles.loc["Outer/Inner"]=np.nan

df_joint_coords_piles.loc[df_joint_coords_piles.XorR==df_joint_coords_piles.XorR.max(),"Outer/Inner"]="Outer"
df_joint_coords_piles.loc[df_joint_coords_piles.XorR==df_joint_coords_piles.XorR.min(),"Outer/Inner"]="Outer"
df_joint_coords_piles.loc[df_joint_coords_piles["Outer/Inner"].isna(),"Outer/Inner"]="Inner"




#%%#%%SET LATERAL SPRINGS
SapModel.SetPresentUnits(6)
# create frame using the coord

left_null_frames=[]
right_null_frames=[]
def create_adjacent_frame(row,left):
    # print(row["XorR"])
    if left: 
        x2=row["XorR"]-0.25
        left_null_frames.append(SapModel.FrameObj.AddByCoord(row["XorR"],0,row["Z"],x2,0,row["Z"],PropName="'NULL BEAM'")[1])
    else:
        x2=row["XorR"]+0.25
        right_null_frames.append(SapModel.FrameObj.AddByCoord(row["XorR"],0,row["Z"],x2,0,row["Z"],PropName="'NULL BEAM'")[1])
        
    

# Create spring support points
# locate the LHS wall and RHS walls points
# get the LHS wall and RHS walls z coord
Joint_coord_left=df_joint_coords.loc[df_joint_coords["XorR"]==df_joint_coords["XorR"].min()]
LHS_joints=Joint_coord_left.copy(deep=True)
LHS_joints[["XorR","Y","Z"]]=LHS_joints[["XorR","Y","Z"]]/1000
print(LHS_joints.apply(lambda x: create_adjacent_frame(x,left=True),axis=1))


Joint_coord_right=df_joint_coords.loc[df_joint_coords["XorR"]==df_joint_coords["XorR"].max()]
RHS_joints=Joint_coord_right.copy(deep=True)
RHS_joints[["XorR","Y","Z"]]=RHS_joints[["XorR","Y","Z"]]/1000
print(RHS_joints.apply(lambda x: create_adjacent_frame(x,left=False),axis=1))

# set t/c limits
for frame in left_null_frames+right_null_frames:
    SapModel.FrameObj.SetTCLimits(frame,False,0,True,0)
    SapModel.FrameObj.SetSection(frame,"NULL BEAM")




# get psudo nodes
Joint_coord2=get_tunnel_sapmodel_info(SapModel)[2]
Joint_coord2[["XorR","Y","Z"]]=Joint_coord2[["XorR","Y","Z"]]/1000
# Joint_coord2["Z"]=(Joint_coord2["Z"]-Joint_coord2["Z"].max())-depth_of_tunnel
Joint_coord_left2=Joint_coord2.loc[Joint_coord2["XorR"]==Joint_coord2["XorR"].min()].sort_values(by="Z",ascending=False).reset_index(drop=True)
Joint_coord_right2=Joint_coord2.loc[Joint_coord2["XorR"]==Joint_coord2["XorR"].max()].sort_values(by="Z",ascending=False).reset_index(drop=True)


df_soil=pd.read_excel(loads_cal_path,sheet_name="Wish In Place",header=39).iloc[:10,2:25].dropna(how="all",axis=0).dropna(how="all",axis=1)
df_soil_profile=pd.read_excel(loads_cal_path,sheet_name="Wish In Place",header=16,usecols="AF:AK").iloc[:25,:]

def Mapping(df,Boundary_Col,Levels_col):
    index_start=int(df[df[Boundary_Col]=='Box Roof CL'].index.values)
    index_end=int(df[df[Boundary_Col]=='Box Base CL'].index.values)
    df=df.iloc[index_start:index_end+1,:]
    df['Levels_Trans']=(df[Levels_col]+(-baseSlab_CL_SHD))
    
    return df

df_soil_profile_mapped=Mapping(df=df_soil_profile,Boundary_Col='Boundary',Levels_col='Level')

df_soil_profile_mapped[['Level', 'Thickness to layer above', "E'", 'kh','Levels_Trans']]=df_soil_profile_mapped[['Level', 'Thickness to layer above', "E'", 'kh','Levels_Trans']].astype("float64")
# find closest value
def find_closest_value(val, df):
    return df.loc[(df['Levels_Trans'] - val).abs().idxmin(),"kh"]

Joint_coord_left2['kh'] = Joint_coord_left2['Z'].apply(lambda x: find_closest_value(x, df_soil_profile_mapped))
Joint_coord_right2['kh'] = Joint_coord_right2['Z'].apply(lambda x: find_closest_value(x, df_soil_profile_mapped))

def get_spring_stiffness(df):
    for index, row in Joint_coord_left2.iterrows():
        if index==0:
            df.loc[index,"stiffness"]=(df.loc[index,'Z']-df.loc[index+1,'Z'])*df.loc[index,'kh']/2
        elif index==len(df)-1:
            df.loc[index,"stiffness"]=(df.loc[index-1,'Z']-df.loc[index,'Z'])*df.loc[index,'kh']/2
        else:
            df.loc[index,"stiffness"]=(df.loc[index-1,'Z']-df.loc[index+1,'Z'])*df.loc[index,'kh']/2
    return df

Joint_coord_left2=get_spring_stiffness(Joint_coord_left2)
Joint_coord_right2=get_spring_stiffness(Joint_coord_right2)

# applying spring to model

print(Joint_coord_left2.apply(lambda row: SapModel.PointObj.SetSpring(Name=row["Joint"],K=[int(row["stiffness"]),0,0,0,0,0],Replace=True),axis=1))
print(Joint_coord_right2.apply(lambda row: SapModel.PointObj.SetSpring(Name=row["Joint"],K=[int(row["stiffness"]),0,0,0,0,0],Replace=True),axis=1))
print("Soil spring set")

SapModel.File.Save(model_path)
print("Model Saved")


#%%#[BOX LOADS] GETTING BOX LOADS
# Userdefined path for load Load Calculation and Spring-Section 

df2=Getting_box_loads5_Ver3.Creating_box_loads(section_type,loads_cal_path,Joint_coord=df_joint_coords,connectivity=df_connectivity) # requires "Joint Coordinates" and "Connectivity - Frame"


#%%[BOX LOADS] APPLYING TO MODEL BOX LOADS
# =============================================================================

# update model loads using df
SapModel.SetPresentUnits(6)

for i in range(len(df2)):
     
    name=df2.loc[i, "Frame"]
    loadPat=df2.loc[i, "LoadPat"]
    myType=1
    Dir=df2.loc[i, "Dir"]
    Dist1=float(0)
    Dist2=float(1)
    Val1=df2.loc[i, "FOverLA"]
    Val2=df2.loc[i, "FOverLB"]
    CSys="GLOBAL"
    RelDist=True
    Replace=True
    ItemType=int(0)

    
    
    SapModel.FrameObj.SetLoadDistributed(name,loadPat,myType,Dir,Dist1,Dist2,Val1,Val2,CSys,RelDist,Replace,ItemType)
    
    i+=1

print("loads updated")

# Apply LC51 wind pressure



left_wall=list(df_connectivity.loc[df_connectivity["Tunnel_component"]=="left_wall"].Frame)
right_wall=list(df_connectivity.loc[df_connectivity["Tunnel_component"]=="right_wall"].Frame)
roof=list(df_connectivity.loc[df_connectivity["Tunnel_component"]=="roof"].Frame)
base_slab=list(df_connectivity.loc[df_connectivity["Tunnel_component"]=="base_slab"].Frame)

list_of_component=[left_wall,right_wall,roof,base_slab]
    
tuple_of_component_and_dir=(
    (left_wall,"left_wall",1),
    (right_wall,"right_wall",-1),
    (roof,"roof",-1),
    (base_slab,"base_slab",1)
    )
    
for component in tuple_of_component_and_dir:
    for frame in component[0]:
        name=frame
        loadPat="51"
        myType=1
        Dir=2
        Dist1=float(0)
        Dist2=float(1)
        Val1=Wind_pressure_internal*component[2] #direction
        Val2=Wind_pressure_internal*component[2] #direction
        CSys="LOCAL"
        RelDist=True
        Replace=True
        ItemType=int(0)
        
        SapModel.FrameObj.SetLoadDistributed(name,loadPat,myType,Dir,Dist1,Dist2,Val1,Val2,CSys,RelDist,Replace,ItemType)
        


# save model
SapModel.File.Save(model_path)
print("Model Saved")


 

#%% 5) RUN ANALYSIS

print("Running analysis")
SapModel.Analyze.RunAnalysis()
print("Analysis completed")

#%% others

# SapObject.ApplicationExit(True)
