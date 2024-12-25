import openpyxl.workbook
import pandas as pd
import os
import glob
import numpy as np
import traceback
from datetime import datetime
from openpyxl import load_workbook
import io
import time
import openpyxl
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
from geopy.distance import geodesic
import pyxlsb
import math
from scipy.spatial import KDTree
import subprocess


output_dir = os.path.join(os.getcwd(), 'OutputFiles')
sites_db = os.path.join(output_dir, 'sites_db.csv')
nbrs_db = os.path.join(output_dir, 'estimated_Nbrs1.csv')
easy_optim_log = os.path.join(output_dir, 'log.xlsx')
para_dump = os.path.join(output_dir, 'dump.xlsb')
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
def audit_Lnadjgnb(Lnadjgnb_audit_form):
    start_time = time.time()
    fileSitesDB_dF = Lnadjgnb_audit_form['fileSitesDB']
    fileParamatersDB_LNADJGNB_DF =   Lnadjgnb_audit_form['fileParamatersDB']
    fileParamatersDB_LNADJGNB_DF.columns = fileParamatersDB_LNADJGNB_DF.columns.str.strip()
    fileSitesDB_dF.columns = fileSitesDB_dF.columns.str.strip()
    fileSitesDB_dF['NodeB'] = fileSitesDB_dF['NodeB'].str.strip()
    fileSitesDB_dF = fileSitesDB_dF.dropna(subset=['NodeB'])
    fileSitesDB_dF = fileSitesDB_dF[fileSitesDB_dF['NodeB'].str.strip() != '']
    columns_latLong = ['NodeB','Long', 'Lat']
    fileSitesDB_dF = fileSitesDB_dF[columns_latLong]
    fileSitesDB_dF['NodeB']= fileSitesDB_dF['NodeB'].apply(pd.to_numeric , errors='coerce')
    fileParamatersDB_LNADJGNB_DF_blankIP = fileParamatersDB_LNADJGNB_DF[
        (fileParamatersDB_LNADJGNB_DF['cPlaneIpAddr'] == "") | (fileParamatersDB_LNADJGNB_DF['cPlaneIpAddr'].isna())]
    fileParamatersDB_LNADJGNB_DF_SitesIP = fileParamatersDB_LNADJGNB_DF[fileParamatersDB_LNADJGNB_DF['adjGnbId'] == fileParamatersDB_LNADJGNB_DF['LNBTS']]
    dict_LNBTSs_IPs = (dict(zip(fileParamatersDB_LNADJGNB_DF_SitesIP['LNBTS'], fileParamatersDB_LNADJGNB_DF_SitesIP['cPlaneIpAddr'])))
    fileParamatersDB_LNADJGNB_DF_blankIP.loc[:, 'cPlaneIpAddr'] = fileParamatersDB_LNADJGNB_DF_blankIP['adjGnbId'].map(dict_LNBTSs_IPs)
    LNADJGNB_Still_Missing_IPs = fileParamatersDB_LNADJGNB_DF_blankIP[
        (fileParamatersDB_LNADJGNB_DF_blankIP['cPlaneIpAddr'].isna())|(fileParamatersDB_LNADJGNB_DF_blankIP['cPlaneIpAddr']=='')]
    fileParamatersDB_LNADJGNB_DF_blankIP = fileParamatersDB_LNADJGNB_DF_blankIP[fileParamatersDB_LNADJGNB_DF_blankIP['cPlaneIpAddr'].notna()]
    dict_longitude = (dict(zip(fileSitesDB_dF['NodeB'], fileSitesDB_dF['Long'])))
    dict_latitude = (dict(zip(fileSitesDB_dF['NodeB'], fileSitesDB_dF['Lat'])))
    fileParamatersDB_LNADJGNB_DF['Source Longitude'] = fileParamatersDB_LNADJGNB_DF['LNBTS'].map(dict_longitude)
    fileParamatersDB_LNADJGNB_DF['Source Latitude'] = fileParamatersDB_LNADJGNB_DF['LNBTS'].map(dict_latitude)
    fileParamatersDB_LNADJGNB_DF['Target Longitude'] = fileParamatersDB_LNADJGNB_DF['adjGnbId'].map(dict_longitude)
    fileParamatersDB_LNADJGNB_DF['Target Latitude'] = fileParamatersDB_LNADJGNB_DF['adjGnbId'].map(dict_latitude)
    fileParamatersDB_LNADJGNB_DF['Distance'] = fileParamatersDB_LNADJGNB_DF.apply(
        lambda row: (
            calculate_distance(
                (row['Source Latitude'], row['Source Longitude']),
                (row['Target Latitude'], row['Target Longitude'])
                ) if pd.notna(row['Source Latitude']) and pd.notna(row['Source Longitude']) 
                and pd.notna(row['Target Latitude']) and pd.notna(row['Target Longitude']) 
                else None
                ), axis=1)
    fileParamatersDB_LNADJGNB_DF = fileParamatersDB_LNADJGNB_DF[fileParamatersDB_LNADJGNB_DF['Distance']>=float(Lnadjgnb_audit_form['no_nbrDelDistance'])]

    LNADJGNB_Audit_output = os.path.join(output_dir, 'LNADJGNB Audit.xlsx')
    LNADJGNB_Audit_Update_XML = os.path.join(output_dir, 'Lnadjgnb_Update.xml')
    LNADJGNB_Audit_Deletete_XML = os.path.join(output_dir, 'Lnadjgnb_Delete.xml')
    make_Update_XML(fileParamatersDB_LNADJGNB_DF_blankIP,'cPlaneIpAddr',LNADJGNB_Audit_Update_XML)
    make_delete_XML(fileParamatersDB_LNADJGNB_DF,LNADJGNB_Audit_Deletete_XML)
    with pd.ExcelWriter(LNADJGNB_Audit_output, engine='openpyxl') as writer:
        fileParamatersDB_LNADJGNB_DF_blankIP.to_excel(writer, sheet_name='Update', index=False)
        fileParamatersDB_LNADJGNB_DF.to_excel(writer, sheet_name='Delete', index=False)
        if len(LNADJGNB_Still_Missing_IPs)>0:
            LNADJGNB_Still_Missing_IPs.to_excel(writer, sheet_name='Still Missing IPs', index=False)
    end_time =time.time()
    duration = str(round((end_time - start_time),0))+" Seconds"
    Lnadjgnb_audit_output = {'duration':duration,
                             'output_Files':[LNADJGNB_Audit_output, LNADJGNB_Audit_Update_XML, LNADJGNB_Audit_Deletete_XML]}
    return Lnadjgnb_audit_output    

def make_Update_XML(tbl_DF, parameter,output_File):
    now = datetime.now()
    formatted_now = now.strftime("%Y-%m-%dT%H:%M:%S")
    root = ET.Element("raml", xmlns="raml21.xsd", version="2.1")
    cm_data = ET.SubElement(root, "cmData", type="plan")
    header = ET.SubElement(cm_data, "header")
    ET.SubElement(header, "log", dateTime=formatted_now, action="created", appInfo="Abdellatif-Ahmed")
    
    for _, row in tbl_DF.iterrows():
        dist_name = f"PLMN-PLMN/MRBTS-{row['MRBTS']}/LNBTS-{row['LNBTS']}/LNADJGNB-{row['LNADJGNB']}"
        managed_object = ET.SubElement(cm_data,"managedObject",class_="LNADJGNB",distName=dist_name,version="xL21A_2012_003",operation="update")
        ET.SubElement(managed_object, "p", name=parameter).text = row[parameter]
    
    xml_data = ET.tostring(root, encoding="utf-8", method="xml")
    with open(output_File, "wb") as xmlfile:
        xmlfile.write(xml_data)

def make_delete_XML(tbl_DF, output_File):
    now = datetime.now()
    formatted_now = now.strftime("%Y-%m-%dT%H:%M:%S")
    root = ET.Element("raml", xmlns="raml21.xsd", version="2.1")
    cm_data = ET.SubElement(root, "cmData", type="plan")
    header = ET.SubElement(cm_data, "header")
    ET.SubElement(header, "log", dateTime=formatted_now, action="created", appInfo="Abdellatif-Ahmed")
    for _, row in tbl_DF.iterrows():
        dist_name = f"PLMN-PLMN/MRBTS-{row['MRBTS']}/LNBTS-{row['LNBTS']}/LNADJGNB-{row['LNADJGNB']}"
        ET.SubElement(
            cm_data, "managedObject", class_="LNADJGNB", distName=dist_name,
            version="xL21A_2012_003", operation="delete"
        )
    xml_data = ET.tostring(root, encoding="utf-8", method="xml")
    
    with open(output_File, "wb") as xmlfile:
        xmlfile.write(xml_data)

def calculate_distance(coord1, coord2):
    return geodesic(coord1, coord2).kilometers

def audit_Lnrel(Lnrel_audit_form):
    print("inside the Tool Function......")
    nbrs_db_df = pd.read_csv(nbrs_db)
    dmp_data_LnCel_DF = Lnrel_audit_form['LnCel']
    dmp_data_LnCel_DF.columns = dmp_data_LnCel_DF.columns.str.strip()
    dmp_data_LnCel_DF['Sector ID'] = dmp_data_LnCel_DF.apply(lambda row: (str(row['LNBTS']) +'_' + str(row['name'])[-2:][:1]), axis=1)
    dmp_data_LnCel_DF['Cell_Lkup'] = dmp_data_LnCel_DF['LNBTS'].astype(str) + '_' + dmp_data_LnCel_DF['LNCEL'].astype(str)
    fileSitesDB_dF = Lnrel_audit_form['fileSitesDB']
    fileParamatersDB_LNREL_DF =   Lnrel_audit_form['fileParamatersDB']
    fileParamatersDB_LNREL_DF.columns = fileParamatersDB_LNREL_DF.columns.str.strip()
    fileSitesDB_dF.columns = fileSitesDB_dF.columns.str.strip()
    fileSitesDB_dF['NodeB'] = fileSitesDB_dF['NodeB'].str.strip()
    fileSitesDB_dF = fileSitesDB_dF.dropna(subset=['NodeB'])
    fileSitesDB_dF = fileSitesDB_dF[fileSitesDB_dF['NodeB'].str.strip() != '']
    columns_latLong = ['NodeB','Long', 'Lat']
    fileSitesDB_dF = fileSitesDB_dF[columns_latLong]
    fileSitesDB_dF['NodeB']= fileSitesDB_dF['NodeB'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF = Lnrel_audit_form['Perform_Data']

    lnrel_Performance_DF['Source Sector ID'] = lnrel_Performance_DF['Adj Intra eNB HO PREP SR'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Adj Intra eNB HO PREP SR'] = lnrel_Performance_DF['Adj Intra eNB HO PREP SR'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Intra eNB HO prep att neigh'] = lnrel_Performance_DF['Intra eNB HO prep att neigh'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Adj Intra eNB HO SR'] = lnrel_Performance_DF['Adj Intra eNB HO SR'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Intra eNB HO attempts per neighbor cell'] = lnrel_Performance_DF['Intra eNB HO attempts per neighbor cell'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Intra eNB HO attempts'] = lnrel_Performance_DF['Intra eNB HO attempts per neighbor cell']
    lnrel_Performance_DF['Intra eNB HO Fails'] = lnrel_Performance_DF['Intra eNB HO attempts per neighbor cell'].astype(float) * (100 - lnrel_Performance_DF['Adj Intra eNB HO SR'].astype(float))

    lnrel_Performance_DF['Number of Inter eNB Handover attempts per neighbor cell relationship'] = lnrel_Performance_DF['Number of Inter eNB Handover attempts per neighbor cell relationship'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Inter eNB HO attempts'] = lnrel_Performance_DF['Number of Inter eNB Handover attempts per neighbor cell relationship']
    lnrel_Performance_DF['Inter eNB NB HO fail ratio'] = lnrel_Performance_DF['Inter eNB NB HO fail ratio'].apply(pd.to_numeric , errors='coerce')
    lnrel_Performance_DF['Inter eNB HO Fails'] = lnrel_Performance_DF['Inter eNB HO attempts'].astype(float) * lnrel_Performance_DF['Inter eNB NB HO fail ratio'].astype(float)

    lnrel_Performance_DF['eci_id'] = lnrel_Performance_DF['eci_id'].apply(pd.to_numeric , errors='coerce')
    
    lnrel_Performance_DF['LNCEL'] = lnrel_Performance_DF['Source LNCEL name'].map(dict(zip(dmp_data_LnCel_DF['name'],dmp_data_LnCel_DF['LNCEL'])))
    lnrel_Performance_DF['LNBTS'] = lnrel_Performance_DF['Source LNCEL name'].map(dict(zip(dmp_data_LnCel_DF['name'],dmp_data_LnCel_DF['LNBTS'])))
    lnrel_Performance_DF['MRBTS'] = lnrel_Performance_DF['Source LNCEL name'].map(dict(zip(dmp_data_LnCel_DF['name'],dmp_data_LnCel_DF['MRBTS'])))
    lnrel_Performance_DF['Source Sector ID'] = lnrel_Performance_DF.apply(lambda row: (str(row['LNBTS']) +'_' + str(row['Source LNCEL name'])[-2:][:1]), axis=1)

    lnrel_Performance_DF['Target Cell'] = lnrel_Performance_DF['eci_id'].apply(get_target_cell_id)
    lnrel_Performance_DF['Target eNB'] = lnrel_Performance_DF['eci_id'].apply(get_target_enb_id)
    lnrel_Performance_DF['Target Lkup'] = lnrel_Performance_DF['Target eNB'].astype(str) + '_' + lnrel_Performance_DF['Target Cell'].astype(str)
    lnrel_Performance_DF['Target Sector ID'] = lnrel_Performance_DF['Target Lkup'].map(dict(zip(dmp_data_LnCel_DF['Cell_Lkup'],dmp_data_LnCel_DF['Sector ID'])))
    lnrel_Performance_DF['FirstTierNeighbors'] = lnrel_Performance_DF['Source Sector ID'].map(dict(zip(nbrs_db_df['Sector_ID'],nbrs_db_df['FirstTierNeighbors'])))
    lnrel_Performance_DF['FirstTierNeighbors'] = lnrel_Performance_DF['FirstTierNeighbors'].fillna('').astype(str)
    lnrel_Performance_DF['Target Sector ID'] = lnrel_Performance_DF['Target Sector ID'].fillna('').astype(str)

    
    
    lnrel_Performance_DF['relation'] = lnrel_Performance_DF['LNBTS'].astype(str) + '_' + lnrel_Performance_DF['LNCEL'].astype(str) + '_' + lnrel_Performance_DF['Target eNB'].astype(str)+ '_' + lnrel_Performance_DF['Target Cell'].astype(str)
    fileParamatersDB_LNREL_DF['relation'] = fileParamatersDB_LNREL_DF['LNBTS'].astype(str) + '_' + fileParamatersDB_LNREL_DF['LNCEL'].astype(str) + '_' + fileParamatersDB_LNREL_DF['ecgiAdjEnbId'].apply(lambda x: str(int(x)) if pd.notnull(x) else '') + '_' + fileParamatersDB_LNREL_DF['ecgiLcrId'].apply(lambda x: str(int(x)) if pd.notnull(x) else '')
    
    lnrel_Performance_DF['LNREL'] = lnrel_Performance_DF['relation'].map(dict(zip(fileParamatersDB_LNREL_DF['relation'],fileParamatersDB_LNREL_DF['LNREL'])))
    lnrel_Performance_DF['cellIndOffNeigh'] = lnrel_Performance_DF['relation'].map(dict(zip(fileParamatersDB_LNREL_DF['relation'],fileParamatersDB_LNREL_DF['cellIndOffNeigh'])))
    lnrel_Performance_DF['handoverAllowed'] = lnrel_Performance_DF['relation'].map(dict(zip(fileParamatersDB_LNREL_DF['relation'],fileParamatersDB_LNREL_DF['handoverAllowed'])))
    lnrel_Performance_DF['removeAllowed'] = lnrel_Performance_DF['relation'].map(dict(zip(fileParamatersDB_LNREL_DF['relation'],fileParamatersDB_LNREL_DF['removeAllowed'])))
    
    dict_longitude = (dict(zip(fileSitesDB_dF['NodeB'], fileSitesDB_dF['Long'])))
    dict_latitude = (dict(zip(fileSitesDB_dF['NodeB'], fileSitesDB_dF['Lat'])))
    lnrel_Performance_DF['Source Longitude'] = lnrel_Performance_DF['LNBTS'].map(dict_longitude)
    lnrel_Performance_DF['Source Latitude'] = lnrel_Performance_DF['LNBTS'].map(dict_latitude)
    lnrel_Performance_DF['Target Longitude'] = lnrel_Performance_DF['Target eNB'].map(dict_longitude)
    lnrel_Performance_DF['Target Latitude'] = lnrel_Performance_DF['Target eNB'].map(dict_latitude)
    lnrel_Performance_DF['Distance'] = lnrel_Performance_DF.apply(
        lambda row: (
            calculate_distance(
                (row['Source Latitude'], row['Source Longitude']),
                (row['Target Latitude'], row['Target Longitude'])
                ) if pd.notna(row['Source Latitude']) and pd.notna(row['Source Longitude']) 
                and pd.notna(row['Target Latitude']) and pd.notna(row['Target Longitude']) 
                else None
                ), axis=1)
    lnrel_Performance_DF['Is_imp_Nbrs'] = lnrel_Performance_DF.apply(lambda r: 'Co-located' if r['Distance'] == 0 else ('yes' if r['Target Sector ID'] in r['FirstTierNeighbors'] else 'no'),axis=1)
    
    out_columns = ['MRBTS','LNBTS','LNCEL','LNREL','Is_imp_Nbrs','Source Sector ID','Target Sector ID','Target eNB','Target Cell','Distance','cellIndOffNeigh','handoverAllowed','removeAllowed','Inter eNB HO attempts','Inter eNB HO Fails','Intra eNB HO attempts','Intra eNB HO Fails']
    output_Table = lnrel_Performance_DF[out_columns]
    # output_Table = lnrel_Performance_DF
    LNREL_Audit_output = os.path.join(output_dir, 'LNREL Audit.xlsx')
    with pd.ExcelWriter(LNREL_Audit_output, engine='openpyxl') as writer:
        output_Table.to_excel(writer, sheet_name='LNREL', index=False)
    return LNREL_Audit_output

def get_target_cell_id(x):
    if pd.notnull(x):
        hex_str = hex(int(x))  # Convert to hex
        # Extract the last 2 digits (cell ID) and convert to decimal
        target_cell_id = int(hex_str[-2:], 16)
        return target_cell_id
    return x

def get_target_enb_id(x):
    if pd.notnull(x):
        hex_str = hex(int(x))  # Convert to hex
        # Extract the rest of the digits except the last 2 (eNB ID) and convert to decimal
        target_enb_id = int(hex_str[:-2], 16) if len(hex_str) > 2 else None
        return target_enb_id
    return x

def clean_Sites_db(sites_df,Is_Update_Nbrs,dB_file_name):
    sites_df.columns = sites_df.columns.str.strip()
    sites_df['NodeB Id'] = sites_df['NodeB Name'].apply(lambda x: x.split('-')[0])
    sites_df['Sector_ID'] = sites_df.apply(lambda row: (str(row['NodeB Id']) +'_' + str(row['Cell Name'])[-2:][:1]), axis=1)
    sites_df["polygon"] = sites_df.apply(lambda row: calculate_sector_polygon(row["Lat"], row["Long"], row["Bore"]), axis=1)
    sites_df["Label"] = sites_df["polygon"].apply(lambda x: x[2])
    sites_df["coordinates"] = sites_df.apply(lambda row: get_coordinate(row["Long"], row["Lat"]), axis=1)
    sites_df.to_csv(sites_db, index=False)
    update_log(dB_file_name,'Sites DB',sites_db)
    if Is_Update_Nbrs:
        get_Nbrs(sites_df,8)
    return sites_df

def calculate_sector_polygon(latitude, longitude, azimuth, radius=0.001):
    # Convert azimuth to radians
    azimuth_rad = math.radians(azimuth)
    width = math.radians(30)  # Sector width of ±15 degrees (30 degrees total)

    # First corner of the triangle
    lat1 = latitude + radius * math.cos(azimuth_rad - width)
    lon1 = longitude + radius * math.sin(azimuth_rad - width)

    # Second corner (main azimuth direction)
    lat2 = latitude + radius * math.cos(azimuth_rad)
    lon2 = longitude + radius * math.sin(azimuth_rad)

    # Third corner of the triangle
    lat3 = latitude + radius * math.cos(azimuth_rad + width)
    lon3 = longitude + radius * math.sin(azimuth_rad + width)

    return [[latitude, longitude], [lat1, lon1], [lat2, lon2], [lat3, lon3], [latitude, longitude]]

def get_coordinate(longitude,latitude):
    return [longitude, latitude]

def get_Nbrs(input_dB,distance):
    selected_col = ['NodeB Id', 'Sector_ID', 'Lat', 'Long', 'Bore']
    sectors_df = input_dB[selected_col].drop_duplicates()
    sectors_df['Azimuth'] = sectors_df['Bore']
    selected_col = ['NodeB Id', 'Lat', 'Long']
    sites_df = input_dB[selected_col].drop_duplicates()
    
    coords = sites_df[['Lat', 'Long']].values
    tree = KDTree(coords)
    sites_df = sites_df.reset_index(drop=True)
    find_neighbors_func = find_neighbors(sites_df, coords, distance, tree,20)
    sites_df['CloseNeighbors'] = sites_df.index.map(find_neighbors_func)
    sectors_df['CloseNeighbors'] = sectors_df['NodeB Id'].map(dict(zip(sites_df['NodeB Id'],sites_df['CloseNeighbors'])))
    sectors_df['FirstTierNeighbors'] = sectors_df.apply(
        lambda row: allocate_neighbors(row, 
                                       row['CloseNeighbors'], 
                                       sectors_df),
        axis=1
    )
    sectors_df.to_csv(nbrs_db, index=False)

def find_neighbors(sites_db, coords, distance_threshold, tree,count):
    def _find(site_idx):
        site_coord = coords[site_idx].reshape(1, -1)
        distances, indices = tree.query(site_coord, k=len(coords), distance_upper_bound=distance_threshold / 111)  # Convert km to degrees approx.
        valid_indices = [i for i, d in zip(indices[0], distances[0]) if i != site_idx and i < len(coords) and d < float('inf')]
        closest_indices = sorted(valid_indices, key=lambda i: distances[0][indices[0].tolist().index(i)])[:count]
        return sites_db.iloc[closest_indices]['NodeB Id'].tolist()

    return _find

def allocate_neighbors(sector, neighbors,sectors_db):
    sector_lat = sector['Lat']
    sector_lon = sector['Long']
    sector_azimuth = sector['Bore']
    
    sector_endpoint1 = endpoint(sector_lat, sector_lon, sector_azimuth, 10)  # 10 km for example
    sector_endpoint2 = endpoint(sector_lat, sector_lon, sector_azimuth+15, 10)
    sector_endpoint3 = endpoint(sector_lat, sector_lon, sector_azimuth-15, 10)
    sector_line1 = (sector_lat, sector_lon), sector_endpoint1
    sector_line2 = (sector_lat, sector_lon), sector_endpoint2
    sector_line3 = (sector_lat, sector_lon), sector_endpoint3

    Sector_Lines = [sector_line1,sector_line2,sector_line3]

    closest_neighbors = []

    for neighbor_site_code in neighbors:
        neighbor_sectors = sectors_db[sectors_db['NodeB Id'] == neighbor_site_code]
        for _, neighbor_sector in neighbor_sectors.iterrows():
            neighbor_lat = neighbor_sector['Lat']
            neighbor_lon = neighbor_sector['Long']
            neighbor_azimuth = neighbor_sector['Bore']
            
            neighbor_endpoint1 = endpoint(neighbor_lat, neighbor_lon, neighbor_azimuth, 10)
            neighbor_endpoint2 = endpoint(neighbor_lat, neighbor_lon, neighbor_azimuth+15, 10)
            neighbor_endpoint3 = endpoint(neighbor_lat, neighbor_lon, neighbor_azimuth-15, 10)
            
            neighbor_line1 = (neighbor_lat, neighbor_lon), neighbor_endpoint1
            neighbor_line2 = (neighbor_lat, neighbor_lon), neighbor_endpoint2
            neighbor_line3 = (neighbor_lat, neighbor_lon), neighbor_endpoint3

            neighbor_lines = [neighbor_line1,neighbor_line2,neighbor_line3]
            
            Sec_Nbr_Lines = []
            for sec_line in Sector_Lines:
                for nbr_line in neighbor_lines:
                    intersect = line_intersection(sec_line[0] , sec_line[1] , nbr_line[0] , nbr_line[1])
                    if intersect:
                        distance_to_intersect = calculate_distance((sector_lat, sector_lon), intersect)
                        Sec_Nbr_Lines.append((distance_to_intersect, neighbor_sector['Sector_ID']))    
            Sec_Nbr_Lines.sort(key=lambda x: x[0])
            if Sec_Nbr_Lines:
                closest_neighbors.append(Sec_Nbr_Lines[0])
            

    closest_neighbors.sort(key=lambda x: x[0])
    return [neighbor[1] for neighbor in closest_neighbors[:20]]

def endpoint(lat, lon, azimuth, distance_km):
    """
    Calculate the endpoint from the start point (lat, lon) given the azimuth and distance.
    """
    R = 6371  # Radius of the Earth in km
    bearing = math.radians(azimuth)
    lat1 = math.radians(lat)
    lon1 = math.radians(lon)
    
    lat2 = math.asin(math.sin(lat1) * math.cos(distance_km / R) +
                     math.cos(lat1) * math.sin(distance_km / R) * math.cos(bearing))
    lon2 = lon1 + math.atan2(math.sin(bearing) * math.sin(distance_km / R) * math.cos(lat1),
                             math.cos(distance_km / R) - math.sin(lat1) * math.sin(lat2))
    
    return (math.degrees(lat2), math.degrees(lon2))

def line_intersection(p1, p2, p3, p4):
    """
    Find the intersection of two lines (p1 -> p2) and (p3 -> p4).
    Returns a tuple (x, y) of the intersection point or None if no intersection.
    """
    s1_x = p2[0] - p1[0]
    s1_y = p2[1] - p1[1]
    s2_x = p4[0] - p3[0]
    s2_y = p4[1] - p3[1]

    denom = -s2_x * s1_y + s1_x * s2_y
    if denom == 0:
        return None  # Lines are parallel

    s = (-s1_y * (p1[0] - p3[0]) + s1_x * (p1[1] - p3[1])) / denom
    t = (s2_x * (p1[1] - p3[1]) - s2_y * (p1[0] - p3[0])) / denom

    if 0 <= s <= 1 and 0 <= t <= 1:
        # Collision detected
        i_x = p1[0] + (t * s1_x)
        i_y = p1[1] + (t * s1_y)
        return (i_x, i_y)

    return None  # No collision

def find_All_Site_neighbors(sites_db, coords, distance_threshold, tree):
    def _find(site_idx):
        site_coord = coords[site_idx].reshape(1, -1)
        distances, indices = tree.query(site_coord, k=len(coords), distance_upper_bound=distance_threshold / 111)  # Convert km to degrees approx.
        valid_indices = [i for i, d in zip(indices[0], distances[0]) if i != site_idx and i < len(coords) and d < float('inf')]
        closest_indices = sorted(valid_indices, key=lambda i: distances[0][indices[0].tolist().index(i)])
        return sites_db.iloc[closest_indices]['NodeB Id'].tolist()

    return _find

def get_log():
    log_df = pd.read_excel(easy_optim_log,sheet_name='log')
    log_df['Upload Date'] = pd.to_datetime(log_df['Upload Date'])
    site_db_files = log_df[log_df['File Type'] == 'Sites DB']
    dmp_files = log_df[log_df['File Type'] == 'Parameters Dump']
    if len(site_db_files)>0:
        most_recent_KML_file = site_db_files.loc[site_db_files['Upload Date'].idxmax()]
        recent_KML_filename = most_recent_KML_file['File Name']
        recent_KML_filelink = most_recent_KML_file['Download Link']
    else:
        recent_KML_filename = None
        recent_KML_filelink = None

    if len(dmp_files)>0:
        most_recent_Dmp_file = dmp_files.loc[dmp_files['Upload Date'].idxmax()]
        recent_Dmp_filename = most_recent_Dmp_file['File Name']
        recent_Dmp_filelink = most_recent_Dmp_file['Download Link']
    else:
        recent_Dmp_filename = None
        recent_Dmp_filelink = None

    return recent_KML_filename,recent_KML_filelink,recent_Dmp_filename,recent_Dmp_filelink

def update_log(dB_file_name, type,file_link):
    file_Date = ""
    if not os.path.exists(easy_optim_log):
        columns = ["File Type", "File Name", "Upload Date", "File Date", "Download Link"]
        log_df = pd.DataFrame(columns=columns)
    else:
        log_df = pd.read_excel(easy_optim_log, sheet_name='log')
    if type == "Sites DB":
        file_Date = dB_file_name[:-4][3:]
    elif type == "Parameters Dump":
        file_Date = dB_file_name[8:]
    else:
        file_Date = "Unknown"

    log_df.loc[len(log_df)] = {
        "File Type": type,
        "File Name": dB_file_name,
        "Upload Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "File Date": dB_file_name[:-4][3:],
        "Download Link":file_link
    }
    with pd.ExcelWriter(easy_optim_log, engine='openpyxl') as writer:
        log_df.to_excel(writer, sheet_name='log', index=False)
def get_log_file():
    return easy_optim_log

def upload_Dmp(file_Dmp):
    try:    
        file_Dmp_Name = file_Dmp.name
        with open(para_dump, "wb") as f:
            f.write(file_Dmp.getbuffer())
        print("file writer successfuly",para_dump)
    except Exception as e:
        print(f"Error processing file: {e}")
    
    update_log(file_Dmp_Name,'Parameters Dump',para_dump)
