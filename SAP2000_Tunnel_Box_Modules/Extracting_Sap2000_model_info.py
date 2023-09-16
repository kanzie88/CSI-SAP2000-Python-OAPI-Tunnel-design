# -*- coding: utf-8 -*-
"""
Created on Fri Aug 25 09:32:15 2023

@author: 10022767
"""


from itertools import zip_longest
import pandas as pd

def get_tunnel_sapmodel_info(SapModel):
    All_joints=SapModel.PointObj.GetNameList()[2]
    All_frames=SapModel.FrameObj.GetNameList()[2]
    
    # for "Connectivity - Frame"
    df_connectivity=pd.DataFrame()
    list_of_joint_coords=[]
    for joint in All_joints:
        result=SapModel.PointObj.GetConnectivity(joint)[3:5]
        a=list(result)
        a.insert(0,tuple([joint]*len(result[1])))
        # list_of_joint_coords.append(a)
        df_connectivity=pd.concat([df_connectivity,pd.DataFrame(list(zip_longest(*a)),columns=["Joint","Frame","PointNumber"],dtype="str")],ignore_index=True)
    
    df_connectivity=df_connectivity.loc[df_connectivity.Frame!=""]
    df_connectivity["JointIJ"]=df_connectivity["PointNumber"].replace({"1":"JointI","2":"JointJ"})
    # 
    
    df_connectivity=df_connectivity.pivot(index="Frame",columns='JointIJ',values='Joint').reset_index()
    df_connectivity=df_connectivity[['Frame', 'JointI','JointJ']]
    
    
    # for "Frame Section Assignments"
    # import the table summary of frames  to consider
    list_of_frame_assignments=[]
    for frame in All_frames:
        list_of_frame_assignments.append([frame,SapModel.FrameObj.GetSection(Name=frame)[1]])
    df_frame_assignment=pd.DataFrame(list_of_frame_assignments,columns=["Frame","AnalSect"])
    df_frames=df_frame_assignment[df_frame_assignment.AnalSect!="NULL BEAM"]
    frame_elements=list(df_frames.Frame.unique())
    
    
    # for "Joint Coordinates"
    list_of_joint_coords=[]
    for joint in All_joints:
        a=list(SapModel.PointObj.GetCoordCartesian(Name=joint)[1:])
        a.insert(0,joint)
        list_of_joint_coords.append(a)
    df_joint_coords=pd.DataFrame(list_of_joint_coords,columns=["Joint","XorR",'Y',"Z"])
    df_joint_coords[["XorR",'Y',"Z"]]=df_joint_coords[["XorR",'Y',"Z"]]*1000
    
    # find inner piles
    
    # dataframe of springs from model
    list_of_spring=[]
    for joint in All_joints:
        spring_resuilt=SapModel.PointObj.GetSpring(joint,[0,1,2,3,4,5])
        if spring_resuilt[0]==0:
            a=list(spring_resuilt[1])
            a.insert(0,joint)
            list_of_spring.append(a)
    
    df_spring_joints=pd.DataFrame(list_of_spring,columns=['Joint','U1','U2','U3','R1','R2','R3'])
    
    df_connectivity=df_connectivity.loc[df_connectivity.Frame.isin(list(df_frames.Frame))]
    # filter out psedo nodes if defined
    if 'Joint_coord_left2' in globals():
       df_joint_coords=df_joint_coords.loc[~df_joint_coords.Joint.isin(list(Joint_coord_left2.Joint))]
    if 'Joint_coord_right2' in globals():   
       df_joint_coords=df_joint_coords.loc[~df_joint_coords.Joint.isin(list(Joint_coord_right2.Joint))]
        
    #getting tunnnel components
    # Locating coordinates from SAP2000 model
    # filtering the left wall
    df_joint_coords_left=df_joint_coords.loc[df_joint_coords["XorR"]==df_joint_coords["XorR"].min()].sort_values(by="Z",ascending=False)
    # filtering the right wall
    df_joint_coords_right=df_joint_coords.loc[df_joint_coords["XorR"]==df_joint_coords["XorR"].max()].sort_values(by="Z",ascending=False)
    # filtering the roof
    df_joint_coords_roof=df_joint_coords.loc[df_joint_coords["Z"]==df_joint_coords["Z"].max()].sort_values(by="Z",ascending=False)
    # filtering the base slab
    df_joint_coords_base=df_joint_coords.loc[df_joint_coords["Z"]==df_joint_coords["Z"].min()].sort_values(by="Z",ascending=False)
   
    def check(df_connectivity,joint_df):
        df_connectivity["JointI_Check"]=df_connectivity["JointI"] in list(joint_df.Joint)
        df_connectivity["JointJ_Check"]=df_connectivity["JointJ"] in list(joint_df.Joint)
       
        return df_connectivity 
   
   #  left wall
    df_connectivity=df_connectivity.apply(lambda x: check(x,df_joint_coords_left),axis=1)
    df_connectivity.loc[df_connectivity.apply(lambda x: (x["JointI_Check"])and(x["JointJ_Check"]),axis=1),"Tunnel_component"]="left_wall"
   # right wall
    df_connectivity=df_connectivity.apply(lambda x: check(x,df_joint_coords_right),axis=1)
    df_connectivity.loc[df_connectivity.apply(lambda x: (x["JointI_Check"])and(x["JointJ_Check"]),axis=1),"Tunnel_component"]="right_wall"
   # roof
    df_connectivity=df_connectivity.apply(lambda x: check(x,df_joint_coords_roof),axis=1)
    df_connectivity.loc[df_connectivity.apply(lambda x: (x["JointI_Check"])and(x["JointJ_Check"]),axis=1),"Tunnel_component"]="roof"
   # base slab
    df_connectivity=df_connectivity.apply(lambda x: check(x,df_joint_coords_base),axis=1)
    df_connectivity.loc[df_connectivity.apply(lambda x: (x["JointI_Check"])and(x["JointJ_Check"]),axis=1),"Tunnel_component"]="base_slab"

    df_connectivity.loc[df_connectivity["Tunnel_component"].isna(),"Tunnel_component"]="internal_wall"
    
    df_connectivity.drop(columns=['JointI_Check', 'JointJ_Check'],inplace=True)
    
    return df_connectivity,df_frames,df_joint_coords,df_spring_joints,frame_elements