# CSI-SAP2000-Python-OAPI-Tunnel-design
INSTRUCTIONS:
Use "SAP2000_Tunnel_Box_modelling.py" for automating your tunnel box creation and design in CSI SAP2000 (customized for LTA CDC 2019 load definations)
Functions of the script:
SAP2000_Tunnel_Box_modelling.py: 
Automated FEM modelling of Box tunnel:
1) defination of load case, pattern, combination definitions,
2) Tunnel box frame creation
3) Meshing of frame element
4) lateral soil spring assignment/pile spring application,
5) calculation of base, roof slab, LHS and RHS walls linearly distributed loads(based on LTA CDC 2019 combnation of action 1,2,3 ,4 for case 4(wish in place structural model)

**replace "Test" with your section "name"**
Steps:
1) Open Axial & Flexural Design-Section Test.xlsm in folder Section Test and update tunnel parameters. Dimenisions of tunnel box, soil profile, size of FEM element,  depth of soil, Live loads, surcharge etc (those in yellow boxes)
2) Update load combinations and load cases in Load_cases_combi.xlsx in sheet [Load summary(updated)] (those in yellow boxes)
<<<<<<< Updated upstream
3) Open SAP2000_Tunnel_Box_modelling.py and run script to create SAP2000 model and analyse
=======
3) Open SAP2000_Tunnel_Box_AnD.py and run script to create to analyze model, and design elements
>>>>>>> Stashed changes


#Use "SAP2000_Tunnel_Box_Modules/Getting_aggreagted_loads.py" to aggregate the four critical cases for RC frame design purpose:
1) 'Max Mz' and correpsonding Fx
2) 'Min Mz' and correpsonding Fx
3) 'Max Fx' and correpsonding Mz
4) 'Min Fx' and correpsonding Mz

 output is a dataframe of aggregated frames forces


![alt text](1694856378451.jpeg)
