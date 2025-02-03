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
import ast
from xml.dom import minidom
from shapely.geometry import Polygon, Point

output_dir = os.path.join(os.getcwd(), 'OutputFiles')
sites_db = os.path.join(output_dir, 'sites_db.csv')
nbrs_db = os.path.join(output_dir, 'estimated_Nbrs1.csv')
easy_optim_log = os.path.join(output_dir, 'log.xlsx')
para_dump = os.path.join(output_dir, 'dump.xlsb')
xml_objects = os.path.join(output_dir, 'XML Objects.xlsx')
created_xml_link = os.path.join(output_dir, 'OutputXML.xml')
xls_PRFILEs = os.path.join(output_dir, 'PRFILE.xlsx')
psc_clash = os.path.join(output_dir, 'Possible Clash Cases.xlsx')
overshooters = os.path.join(output_dir, 'Overshooting Sectors.xlsx')
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
    dom = minidom.parseString(xml_data)  # Parse the XML string
    pretty_xml = dom.toprettyxml(indent="  ")  # Add indentation and line breaks
    with open(output_File, "w", encoding="utf-8") as xmlfile:
                xmlfile.write(pretty_xml)

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
    dom = minidom.parseString(xml_data)  # Parse the XML string
    pretty_xml = dom.toprettyxml(indent="  ")  # Add indentation and line breaks
    with open(output_File, "w", encoding="utf-8") as xmlfile:
                xmlfile.write(pretty_xml)


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

    fileSitesDB_dF_unique1 = fileSitesDB_dF[['NodeB', 'Long']].drop_duplicates(subset=['NodeB'])
    dict_longitude = dict(zip(fileSitesDB_dF_unique1['NodeB'], fileSitesDB_dF_unique1['Long']))
    fileSitesDB_dF_unique2 = fileSitesDB_dF[['NodeB', 'Lat']].drop_duplicates(subset=['NodeB'])
    dict_latitude = dict(zip(fileSitesDB_dF_unique2['NodeB'], fileSitesDB_dF_unique2['Lat']))
    
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
    sites_df['NodeB Id'] = sites_df['NodeB Name'].apply(lambda x: str(x).split('-')[0])
    sites_df['Sector_ID'] = sites_df.apply(lambda row: (str(row['NodeB Id']) +'_' + str(row['Cell Name'])[-2:][:1]), axis=1)
    sites_df['Lat']  = pd.to_numeric(sites_df['Lat'] , errors='coerce')
    sites_df['Long']  = pd.to_numeric(sites_df['Long'] , errors='coerce')
    sites_df['Bore']  = pd.to_numeric(sites_df['Bore'] , errors='coerce')
    required_columns = ['Long', 'Lat', 'Bore']
    sites_df = sites_df.dropna(subset=required_columns)
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
        recent_KML_filename = ""
        recent_KML_filelink = ""

    if len(dmp_files)>0:
        most_recent_Dmp_file = dmp_files.loc[dmp_files['Upload Date'].idxmax()]
        recent_Dmp_filename = most_recent_Dmp_file['File Name']
        recent_Dmp_filelink = most_recent_Dmp_file['Download Link']
    else:
        recent_Dmp_filename = ""
        recent_Dmp_filelink = ""

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
        update_log(file_Dmp_Name,'Parameters Dump',para_dump)
        print("file writer successfuly",para_dump)
    except Exception as e:
        print(f"Error processing file: {e}")
    
    update_log(file_Dmp_Name,'Parameters Dump',para_dump)

def valide_make_XML(selected_Object, changes_csv,action):
    df_xml_objects = pd.read_excel(xml_objects)
    filtered_row = df_xml_objects[df_xml_objects["Object"] == selected_Object].iloc[0]
    mandatory_columns = filtered_row["Mandatory_ID"]
    mandatory_columns = ast.literal_eval(mandatory_columns)
    parameters_list = filtered_row["Parameters_List"]
    parameters_list = ast.literal_eval(parameters_list)
    df_changes_csv = pd.read_csv(changes_csv)
    # validate Mandatory Columns
    output_string = []
    missing_columns = []
    unrelated_parameters = []
    output_string.append(mandatory_columns)
    output_string.append(df_changes_csv.columns)
    for column in mandatory_columns:
        if column not in df_changes_csv.columns:
            missing_columns.append(column)

    for col in df_changes_csv.columns :
        if col not in mandatory_columns: 
            if col not in parameters_list:
                unrelated_parameters.append(col)
    
    if len(missing_columns) > 0:
        return "Missing Mandatory Identifiers", missing_columns
    else:
        if len(unrelated_parameters) > 0:
            return "given Parameters to change >>> " + str(unrelated_parameters) +  " <<< are not in such object  " + str(selected_Object)
        else:
            # Now everything is ok, we will build the XML
            now = datetime.now()
            formatted_now = now.strftime("%Y-%m-%dT%H:%M:%S")
            root = ET.Element("raml", xmlns="raml21.xsd", version="2.1")
            cm_data = ET.SubElement(root, "cmData", type="plan")
            header = ET.SubElement(cm_data, "header")
            ET.SubElement(header, "log", dateTime=formatted_now, action="created", appInfo="Abdellatif-Ahmed")
            for _, row in df_changes_csv.iterrows():
                dist_name = "PLMN-PLMN/"
                i= 0
                for col in mandatory_columns:
                    if i < len(mandatory_columns)-1:
                        dist_name = dist_name + col + "-" + str(row[col]) + "/"
                    else:
                        dist_name = dist_name + col + "-" + str(row[col])
                    i = i+1
                # if action == "update" or action == "create":
                #     managed_object = ET.SubElement(cm_data,"managedObject",class_=selected_Object,distName=dist_name,version="xL21A_2012_003",operation=action)
                #     for parameter in df_changes_csv.columns:
                #         if parameter not in mandatory_columns:
                #             ET.SubElement(managed_object, "p", name=parameter).text = str(row[parameter])
                # elif action == "delete":
                #     ET.SubElement(cm_data, "managedObject", class_=selected_Object, distName=dist_name,version="xL21A_2012_003", operation=action)
                # else:
                #     return "Error Un-Identified operation!"
                if action == "update" or action == "create":
                    managed_object = ET.SubElement(cm_data,"managedObject",**{"class": selected_Object, "distName": dist_name, "version": "xL21A_2012_003", "operation": action})
                    for parameter in df_changes_csv.columns:
                        if parameter not in mandatory_columns:
                            ET.SubElement(managed_object, "p", name=parameter).text = str(row[parameter])
                elif action == "delete":
                    ET.SubElement(cm_data,"managedObject",**{"class": selected_Object, "distName": dist_name, "version": "xL21A_2012_003", "operation": action})
                else:
                    return "Error Un-Identified operation!"

            xml_data = ET.tostring(root, encoding="utf-8", method="xml")
            dom = minidom.parseString(xml_data)  # Parse the XML string
            pretty_xml = dom.toprettyxml(indent="  ")  # Add indentation and line breaks
            with open(created_xml_link, "w", encoding="utf-8") as xmlfile:
                xmlfile.write(pretty_xml)  
            # with open(created_xml_link, "wb") as xmlfile:
            #     xmlfile.write(xml_data)
            return "XML Created Successfully!"
            
def audit_prfiles(files):
    try:
        print("inside the function")
        data = []
        for prfile in files:
            prfile_name = prfile.name[:-4]
            lines = prfile.getvalue().decode("utf-8").splitlines()
            current_class = None
            for line in lines:
                line = line.strip()
                if line.startswith("PARAMETER CLASS:"):
                    # Extract the class name
                    current_class = line.replace("PARAMETER CLASS:", "").strip()
                elif line and line[0].isdigit():
                    # Extract parameter data
                    parts = line.split()
                    identifier = parts[0]
                    parameter_name = " ".join(parts[1:-2])
                    value_hex = parts[-2]
                    change_possibility = parts[-1]
                    data.append([prfile_name,current_class, identifier, parameter_name, value_hex, change_possibility])
        df_Prfile = pd.DataFrame(data, columns=["RNC", "Parameter Class", "Identifier", "Parameter Name", "Value (Hex)", "Change Possibility"])
        pivot_df = df_Prfile.pivot_table(
            index=["Parameter Class", "Identifier", "Parameter Name", "Change Possibility"],
            columns="RNC",
            values=["Value (Hex)"],
            aggfunc="first"  # Since there should be one value per combination
        )
        pivot_df.columns = [f"{col[1]} - {col[0]}" for col in pivot_df.columns]
        pivot_df.reset_index(inplace=True)
        with pd.ExcelWriter(xls_PRFILEs, engine='openpyxl') as writer:
            pivot_df.to_excel(writer, sheet_name='PRFILE', index=False)
        return "PRFILEs Preparation Done Successfully!"
    except Exception as e:
        return e
def get_wcl(df_criteria,wcl_KPIs,tech,thrshld_days):
    start_time = time.time()
    WCL_file = os.path.join(output_dir, tech +'_WCL.xlsx')
    if len(wcl_KPIs)>0:
        print(tech," has ",len(wcl_KPIs)," files")
        WCL_Criteria_url = os.path.join(output_dir, 'WCL_Criteria.xlsx')
        identity_columns = []
        date_columns = []
        patterns = [ 'site', 'wbts', 'nodeb', 'node b','wcel', 'lncel', 'lnbts' ]
        date_patterns = ['date', 'time','period']
        date_col = ''
        i=0
        kpis_csvs = []
        for kpis_file in wcl_KPIs:
            try:
                kpis_df = pd.read_csv(kpis_file, delimiter=';')
                kpis_csvs.append(kpis_df)
            except Exception as e:
                print(f"Error reading file {kpis_file}: {e}")          
            identity_cols = [col for col in kpis_df.columns if any(pattern in col.lower() for pattern in patterns)]
            # identity_cols = [value for value in identity_cols if 'asiacell' not in value.lower()]
            date_cols = [col for col in kpis_df.columns if any(pattern in col.lower() for pattern in date_patterns)]
            if i==0:
                identity_columns = identity_cols
                date_columns = date_cols
            else:
                identity_columns = [value for value in identity_cols if value in identity_columns]
                date_columns = [value for value in date_cols if value in date_columns]

            i = i+1
        date_col = date_columns[0]
        # df = pd.read_excel(WCL_Criteria_url)
        criteria = df_criteria[df_criteria['Technology']==tech]
        criteria = criteria.fillna('')
        not_exiting_Kpis = []
        output = {}
        output_summary = {}
        for index, row in criteria.iterrows():
            try:
                kpis_cols = identity_columns.copy()
                kpis_cols.append(date_col)
                # print(kpis_cols)
                kpi_name = str(row['KPI_Name'])
                kpis_cols.append(str(row['Indicator1']))
                if str(row['Indicator2']) != '':
                    kpis_cols.append(str(row['Indicator2']))
                selected_Kpi = pd.DataFrame()
                missing_indicators = []
                for kpis_file in kpis_csvs:
                    if str(row['Indicator1']) in kpis_file.columns:
                        if str(row['Indicator2']) != '' and str(row['Indicator2']) in kpis_file.columns:
                            selected_Kpi = pd.concat([selected_Kpi,kpis_file[kpis_cols]],ignore_index = True)
                        elif str(row['Indicator2']) == '':
                            selected_Kpi = pd.concat([selected_Kpi,kpis_file[kpis_cols]],ignore_index = True)
                        else:
                            missing_indicators.append(str(row['Indicator2']))
                    else:
                        missing_indicators.append(str(row['Indicator1']))

                selected_Kpi.drop_duplicates(inplace=True)
                try:
                    selected_Kpi[str(row['Indicator1'])] = selected_Kpi[str(row['Indicator1'])].apply(pd.to_numeric , errors='coerce')
                    selected_Kpi['TH1_Crossed'] = apply_condition(selected_Kpi[str(row['Indicator1'])], str(row['Logical_Condition1']) , float(str(row['Threshold1'])))
                except:
                    
                    print(kpi_name,":first indicator had a problem")
                
                if str(row['Indicator2']) != '': 
                    selected_Kpi[str(row['Indicator2'])] = selected_Kpi[str(row['Indicator2'])].apply(pd.to_numeric , errors='coerce')
                    selected_Kpi['TH2_Crossed'] = apply_condition(selected_Kpi[str(row['Indicator2'])], str(row['Logical_Condition2']) , float(str(row['Threshold2'])))
                    selected_Kpi['THLDs_Crossed'] = (selected_Kpi['TH1_Crossed'] & selected_Kpi['TH2_Crossed']).astype(int)
                else:
                    selected_Kpi['THLDs_Crossed'] = (selected_Kpi['TH1_Crossed']).astype(int)
                
                avail_dates = len(selected_Kpi[date_col].unique())
                if avail_dates >= thrshld_days:
                    if str(row['Indicator2']) != '':
                        kpi_pivot = selected_Kpi.pivot_table(values=[str(row['Indicator1']),str(row['Indicator2']),'THLDs_Crossed'], index=identity_columns, columns=date_col, aggfunc='sum', fill_value=0)
                    else:
                        kpi_pivot = selected_Kpi.pivot_table(values=[str(row['Indicator1']),'THLDs_Crossed'], index=identity_columns, columns=date_col, aggfunc='sum', fill_value=0)
                    kpi_pivot.reset_index(inplace=True)
                    kpi_pivot.columns = ['_'.join(map(str, col)).strip() for col in kpi_pivot.columns.values]
                    thrshld_col = [value for value in kpi_pivot.columns if 'THLDs_Crossed' in value]
                    kpi_pivot['Problematic'] = kpi_pivot[thrshld_col].sum(axis=1)
                    kpi_pivot = kpi_pivot[kpi_pivot['Problematic']>=thrshld_days]

                    kpi_pivot = kpi_pivot.drop(columns=thrshld_col)
                    # print(kpi_name)
                    # print(kpi_pivot)
                    output[kpi_name] = kpi_pivot
                    output_summary[kpi_name] = len(kpi_pivot)
            except Exception as e:
                not_exiting_Kpis.append(kpi_name)
                output_summary[kpi_name] = "Error: Missing Counters/Indicators" 
        
        if output_summary:
            summary = pd.DataFrame.from_dict(output_summary, orient='index', columns=['Count'])
            print(summary)
        else:
            print("No KPIs met the criteria.")
        if len(output_summary)>0:
            with pd.ExcelWriter(WCL_file, engine='openpyxl') as writer:
                summary.to_excel(writer, sheet_name='Summary', index=True)
            with pd.ExcelWriter(WCL_file, engine='openpyxl',mode='a') as writer:
                for kpi, wcls in output.items():
                    wcls.to_excel(writer, sheet_name=kpi, index=True)
            end_time =time.time()
            duration = str(round((end_time - start_time),0))+" Seconds"
            print(duration)
            return WCL_file
        else:
            print("missing Counters or KPIs in the given files")
    else:
        print("no KPIs received in ", tech)

def apply_condition(indicator, condition, threshold):
    if condition == '>':
        return indicator > threshold
    elif condition == '<':
        return indicator < threshold
    elif condition == '=':
        return indicator == threshold
    elif condition == '>=':
        return indicator >= threshold
    elif condition == '<=':
        return indicator <= threshold
    else:
        raise ValueError(f"Unknown condition: {condition}")    
    
def checkPSC_Nbrs_Clash(kmlFile,dumpFile):
    start_time = time.time()
    print("PSC Clashes tool Started!")
    # df_kml = pd.read_csv(kmlFile, engine='python', encoding='Windows-1252')
    df_kml = pd.read_excel(kmlFile)
    df_kml.columns = df_kml.columns.str.strip()
    df_kml['NodeB Id'] = df_kml['NodeB Name'].apply(lambda x: x.split('-')[0])
    df_kml['Sector_ID'] = df_kml.apply(lambda row: (str(row['NodeB Id']) +'_' + str(row['Cell Name'])[-2:][:1]), axis=1)
    xls = pd.ExcelFile(dumpFile, engine='pyxlsb')
    df_wcel = pd.read_excel(xls, sheet_name='WCEL', header=1)
    df_wcel.columns = df_wcel.columns.str.strip() 
    df_wcel['cell_Lkup'] = df_wcel['RNC'].astype(str) + '_' + df_wcel['WBTS'].astype(str) + '_' + df_wcel['WCEL'].astype(str)
    df_wcel['cell_Lkup2'] = df_wcel['RNC'].astype(str) + '_' + df_wcel['WCEL'].astype(str)
    df_wcel['Sector ID'] = df_wcel.apply(lambda row: (str(row['WBTS']) +'_' + str(row['name'])[-2:][:1]), axis=1)
    df_wcel['Site Latitude'] = df_wcel['Sector ID'].map(dict(zip(df_kml['Sector_ID'],df_kml['Lat'])))
    df_wcel['Site Longitude'] = df_wcel['Sector ID'].map(dict(zip(df_kml['Sector_ID'],df_kml['Long'])))

    df_adjs = pd.read_excel(xls, sheet_name='ADJS', header=1)
    df_adjs.columns = df_adjs.columns.str.strip()
    df_adjs['cell_Lkup'] = df_adjs['RNC'].astype(str) + '_' + df_adjs['WBTS'].astype(str) + '_' + df_adjs['WCEL'].astype(str)
    df_adjs['Source Cell Name'] = df_adjs['cell_Lkup'].map(dict(zip(df_wcel['cell_Lkup'],df_wcel['name'])))
    df_adjs['Source Sector ID'] = df_adjs['cell_Lkup'].map(dict(zip(df_wcel['cell_Lkup'],df_wcel['Sector ID'])))
    df_adjs['Source PSC'] = df_adjs['cell_Lkup'].map(dict(zip(df_wcel['cell_Lkup'],df_wcel['PriScrCode'])))
    df_adjs['Source Carrier'] = df_adjs['cell_Lkup'].map(dict(zip(df_wcel['cell_Lkup'],df_wcel['UARFCN'])))
    df_adjs['Source Latitude'] = df_adjs['Source Sector ID'].map(dict(zip(df_kml['Sector_ID'],df_kml['Lat'])))
    df_adjs['Source Longitude'] = df_adjs['Source Sector ID'].map(dict(zip(df_kml['Sector_ID'],df_kml['Long'])))

    df_adjs['Target Lkup'] = df_adjs['AdjsRNCid'].astype(str) + '_' + df_adjs['AdjsCI'].astype(str)
    df_adjs['Target Cell Name'] = df_adjs['Target Lkup'].map(dict(zip(df_wcel['cell_Lkup2'],df_wcel['name'])))
    # df_adjs['Target Cell Name'] = df_adjs['name']
    df_adjs['Target Sector ID'] = df_adjs.apply(lambda row: (str(row['Target Cell Name'])[:4] +'_' + str(row['Target Cell Name'])[-2:][:1]), axis=1)
    df_adjs['Target PSC'] = df_adjs['AdjsScrCode']
    df_adjs['Target Latitude'] = df_adjs['Target Sector ID'].map(dict(zip(df_kml['Sector_ID'],df_kml['Lat'])))
    df_adjs['Target Longitude'] = df_adjs['Target Sector ID'].map(dict(zip(df_kml['Sector_ID'],df_kml['Long'])))
    df_adjs['Distance'] = df_adjs.apply(
        lambda row: (
            calculate_distance(
                (row['Source Latitude'], row['Source Longitude']),
                (row['Target Latitude'], row['Target Longitude'])
                ) if pd.notna(row['Source Latitude']) and pd.notna(row['Source Longitude']) 
                and pd.notna(row['Target Latitude']) and pd.notna(row['Target Longitude']) 
                else None
                ), axis=1)
    print("Now going to get the CLosest Cell with Same PSC as the Target Cell!")
    df_adjs['Possible Clash Cell'], df_adjs['Possible PSC Clash Distance'] = zip(*df_adjs.apply(lambda row: calculate_shortest_distance(row, df_kml),axis=1))
    df_adjs['Possible PSC Clash (Y/N)'] = df_adjs.apply(lambda row: "Y" if row['Possible PSC Clash Distance'] <= row['Distance'] else "N",axis=1)
    df_output_possible_clashes = df_adjs[df_adjs['Possible PSC Clash (Y/N)'] == "Y"]
    df_output_possible_clashes = df_output_possible_clashes [df_output_possible_clashes ['Possible Clash Cell']!= df_output_possible_clashes ['Target Cell Name']]
    
    with pd.ExcelWriter(psc_clash, engine='openpyxl') as writer:
        df_output_possible_clashes.to_excel(writer, sheet_name='Possible Clashes1', index=False)
    
    end_time =time.time()
    duration = str(round((end_time - start_time),0))+" Seconds"
    return duration

def calculate_shortest_distance(row, df_kml):
    target_psc = row['Target PSC']
    source_carrier = row['Source Carrier']
    target_sector = row['Target Sector ID']
    source_lat, source_lon = row['Source Latitude'], row['Source Longitude']
    
    # Check if source coordinates are valid
    if pd.notna(source_lat) and pd.notna(source_lon) and pd.notna(target_psc):
        # Filter cells in df_kml with the same PriScrCode as Target PSC
        filtered_cells = df_kml[df_kml['DL Primary Scrambling Code'] == target_psc]
        filtered_cells = filtered_cells[filtered_cells['Downlink UARFCN'] == source_carrier]
        filtered_cells = filtered_cells[filtered_cells['Sector_ID'] != target_sector]  # Ensure it's not the same sector
        # Remove rows with invalid coordinates from filtered_cells
        filtered_cells = filtered_cells.dropna(subset=['Lat', 'Long'])
        
        if not filtered_cells.empty:
            # Calculate distances to each filtered cell
            distances = filtered_cells.apply(
                lambda w_row: calculate_distance(
                    (source_lat, source_lon),
                    (w_row['Lat'], w_row['Long'])
                ),
                axis=1
            )
            # Get the minimum distance and corresponding cell name
            if not distances.empty:
                min_distance_index = distances.idxmin()
                min_distance = distances[min_distance_index]
                clash_cell = filtered_cells.loc[min_distance_index, 'Cell Name']
                return clash_cell, min_distance
    
    return None, None

def get_overshooters(tech, pd_report, KML_DB):
    start_time = time.time()
    if tech =="3G":
        print("Selected to get overshooters for 3G Technology!")
        df_propag_report = pd.read_excel(pd_report)
        df_kml = pd.read_excel(KML_DB)
        df_kml.columns = df_kml.columns.str.strip()
        report_Nodes = df_propag_report.rename(columns={"WBTS ID": "NodeB Id"}).iloc[1:]
        report_Nodes = report_Nodes['NodeB Id'].unique()
        df_kml['NodeB Id'] = df_kml['NodeB']
        selected_col = ['NodeB Id', 'Lat', 'Long']
        df_sites = df_kml[selected_col].drop_duplicates()
        coords = df_sites[['Lat', 'Long']].values
        tree = KDTree(coords)
        df_sites = df_sites.reset_index(drop=True)
        find_neighbors_func = find_neighbors2(df_sites, coords, 10, tree,1000,report_Nodes)
        df_sites['CloseNeighbors'] = df_sites.index.map(find_neighbors_func)
        df_propag_report = df_propag_report.iloc[1:]
        date_patterns = ['date', 'time','period']
        date_cols = [col for col in df_propag_report.columns if any(pattern in col.lower() for pattern in date_patterns)]
        patterns = ['rnc', 'site', 'wbts', 'nodeb', 'node b','wcel' ]
        identity_cols = [col for col in df_propag_report.columns if any(pattern in col.lower() for pattern in patterns)]
        pd_col_patterns = ['propagation_delay', 'prach_delay_average']
        prop_cols = [col for col in df_propag_report.columns if any(pattern in col.lower() for pattern in pd_col_patterns)]
        most_recent_data = max(df_propag_report[date_cols[0]].unique())
        df_filtered = df_propag_report[df_propag_report[date_cols[0]] == most_recent_data]
        needed_cols = identity_cols.copy()
        pd_col = prop_cols[0]
        needed_cols.append(pd_col)

        df_filtered = df_filtered[needed_cols]
        df_filtered['Lat'] = df_filtered['WCEL name'].map(dict(zip(df_kml['Cell Name'],df_kml['Lat'])))
        df_filtered['Long'] = df_filtered['WCEL name'].map(dict(zip(df_kml['Cell Name'],df_kml['Long'])))
        df_filtered['Azimuth'] = df_filtered['WCEL name'].map(dict(zip(df_kml['Cell Name'],df_kml['Bore'])))
        df_filtered['Azimuth'] = df_filtered['Azimuth'].apply(pd.to_numeric , errors='coerce')
        df_filtered['Propagation_polygon'] = df_filtered.apply(lambda row: calculate_polygon(row['Lat'],row['Long'],row['Azimuth'],row[pd_col],Antenna_Beamwidth=50),axis=1)
        df_filtered['CloseNeighbors'] = df_filtered['WBTS ID'].map(dict(zip(df_sites['NodeB Id'],df_sites['CloseNeighbors'])))
        df_filtered['Overshooted_Sites'] = df_filtered.apply(lambda row: count_points_inside_polygon(row,df_sites,'WCEL name'),axis=1)
        non_existing_cells = df_filtered[df_filtered['Lat'].isna()]
        df_filtered= df_filtered[df_filtered['Overshooted_Sites']>0]
        df_filtered = df_filtered.drop(columns=['Propagation_polygon', 'CloseNeighbors'])
        
        with pd.ExcelWriter(overshooters, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, sheet_name='3G_Overshooting Sectors', index=False)
            non_existing_cells.to_excel(writer,sheet_name='Missing 3G Cells in Site DB', index=False) 
        end_time =time.time()
        duration = str(round((end_time - start_time),0))+" Seconds"
        return duration

    if tech =="4G":
        print("Selected to get overshooters for LTE Technology!")
        df_propag_report = pd.read_excel(pd_report)
        report_Nodes = df_propag_report.iloc[1:]
        report_Nodes['NodeB Id'] = report_Nodes['LNBTS name'].str.extract(r'(\d+)-').astype(float)
        report_Nodes= report_Nodes['NodeB Id'].unique()
        df_kml = pd.read_excel(KML_DB)
        df_kml.columns = df_kml.columns.str.strip()
        df_kml['NodeB Id'] = df_kml['eNodeB ID']
        selected_col = ['NodeB Id', 'Lat', 'Long']
        df_sites = df_kml[selected_col].drop_duplicates()
        coords = df_sites[['Lat', 'Long']].values
        tree = KDTree(coords)
        df_sites = df_sites.reset_index(drop=True)
        find_neighbors_func = find_neighbors2(df_sites, coords, 10, tree,1000,report_Nodes)
        df_sites['CloseNeighbors'] = df_sites.index.map(find_neighbors_func)
        df_propag_report = df_propag_report.iloc[1:]
        date_patterns = ['date', 'time','period']
        date_cols = [col for col in df_propag_report.columns if any(pattern in col.lower() for pattern in date_patterns)]
        patterns = ['mrbts name', 'lnbts name', 'lncel name' ]
        identity_cols = [col for col in df_propag_report.columns if any(pattern in col.lower() for pattern in patterns)]
        
        most_recent_data = max(df_propag_report[date_cols[0]].unique())
        df_filtered = df_propag_report[df_propag_report[date_cols[0]] == most_recent_data]
        needed_cols = identity_cols.copy()
        pd_col_patterns = ['avg ue distance']
        prop_cols = [col for col in df_propag_report.columns if any(pattern in col.lower() for pattern in pd_col_patterns)]
        pd_col = prop_cols[0]
        # print(pd_col)
        needed_cols.append(pd_col)
        df_filtered = df_filtered[needed_cols]
        df_filtered[pd_col] = df_filtered[pd_col]*1000
        df_filtered['LNBTS ID'] = df_filtered['LNBTS name'].str.extract(r'(\d+)-').astype(float)
        df_filtered['Lat'] = df_filtered['LNCEL name'].map(dict(zip(df_kml['Cell Name'],df_kml['Lat'])))
        df_filtered['Long'] = df_filtered['LNCEL name'].map(dict(zip(df_kml['Cell Name'],df_kml['Long'])))
        df_filtered['Azimuth'] = df_filtered['LNCEL name'].map(dict(zip(df_kml['Cell Name'],df_kml['Azimuth'])))
        df_filtered['Azimuth'] = df_filtered['Azimuth'].apply(pd.to_numeric , errors='coerce')
        df_filtered['Propagation_polygon'] = df_filtered.apply(lambda row: calculate_polygon(row['Lat'],row['Long'],row['Azimuth'],row[pd_col],Antenna_Beamwidth=50),axis=1)
        df_filtered['CloseNeighbors'] = df_filtered['LNBTS ID'].map(dict(zip(df_sites['NodeB Id'],df_sites['CloseNeighbors'])))
        df_filtered['Overshooted_Sites'] = df_filtered.apply(lambda row: count_points_inside_polygon(row,df_sites,'LNCEL name'),axis=1)
        non_existing_cells = df_filtered[df_filtered['Lat'].isna()]
        df_filtered= df_filtered[df_filtered['Overshooted_Sites']>0]
        df_filtered = df_filtered.drop(columns=['Propagation_polygon', 'CloseNeighbors'])

        with pd.ExcelWriter(overshooters, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, sheet_name='4G_Overshooting Sectors', index=False)
            non_existing_cells.to_excel(writer,sheet_name='Missing 4G Cells in Site DB', index=False)
        end_time =time.time()
        duration = str(round((end_time - start_time),0))+" Seconds"
        return duration

def count_points_inside_polygon(row, df2,cell_name):
    polygon_points = row['Propagation_polygon']
    
    count = 0
    checked_Nbrs = 0
    existing_Nbrs = 0
    
    # if pd.isna(row['AllNeighbors']):
    #     return count
    
    neighbors_list = str(row['CloseNeighbors']).replace("'", "").replace("[", "").replace("]", "").replace(" ", "").split(',')
    neighbors_list = [nbr.strip() for nbr in neighbors_list if nbr]
    # print(neighbors_list)
    existing_Nbrs = len(neighbors_list)

    # df2_filtered = df2[df2['NodeB Id'].isin(neighbors_list)]
    df2_filtered = df2[df2['NodeB Id'].astype(str).isin(neighbors_list)]
    checked_Nbrs = len(df2_filtered)
    for index, site_row in df2_filtered.iterrows():
        site_point = (site_row['Lat'], site_row['Long'])
        if is_point_inside_polygon(polygon_points, site_point):
            count += 1
    # print(checked_Nbrs,":",existing_Nbrs,":",row[cell_name],":",count)
    return count
    
def is_point_inside_polygon(polygon_points, point):
    polygon = Polygon(polygon_points)
    point = Point(point)
    return polygon.contains(point)

def find_neighbors2(sites_db, coords, distance_threshold, tree, count, nodes):
    # Convert `nodes` to a set for efficient lookups
    nodes_set = set(nodes)
    def _find(site_idx):
        # Get the NodeB Id for the current site
        nodeb_id = sites_db.iloc[site_idx]['NodeB Id']
        # Check if the current NodeB Id is in the nodes set
        if nodeb_id in nodes_set:
            site_coord = coords[site_idx].reshape(1, -1)
            distances, indices = tree.query(
                site_coord, k=len(coords), distance_upper_bound=distance_threshold / 111
            )  # Convert km to degrees approx.
            valid_indices = [
                i for i, d in zip(indices[0], distances[0])
                if i != site_idx and i < len(coords) and d < float('inf')
            ]
            closest_indices = sorted(
                valid_indices, key=lambda i: distances[0][indices[0].tolist().index(i)]
            )[:count]
            return sites_db.iloc[closest_indices]['NodeB Id'].tolist()
        else:
            return "Not Needed"
    return _find

def calculate_polygon(latitude, longitude, azimuth, radius_meters=100,Antenna_Beamwidth =30):
    """
    Calculate the polygon for a sector based on latitude, longitude, azimuth, and radius in meters.

    Args:
        latitude (float): Latitude of the center point.
        longitude (float): Longitude of the center point.
        azimuth (float): Direction of the sector in degrees.
        radius_meters (float): Radius of the sector in meters (default is 100 meters).

    Returns:
        list: A list of [latitude, longitude] pairs defining the sector polygon.
    """
    if pd.isnull(latitude) or pd.isnull(longitude) or pd.isnull(azimuth) or pd.isnull(radius_meters):
        return None 
    # Earth's radius in meters
    earth_radius = 6378137  # WGS84

    # Convert azimuth and width to radians
    azimuth_rad = math.radians(azimuth)
    width = math.radians(Antenna_Beamwidth)  # Sector width of ±15 degrees (30 degrees total)

    # Conversion factors for meters to latitude and longitude degrees
    radius_lat = radius_meters / earth_radius  # Convert radius to radians latitude
    radius_lon = radius_meters / (earth_radius * math.cos(math.radians(latitude)))  # Adjust for longitude compression

    # First corner of the triangle
    lat1 = latitude + math.degrees(radius_lat * math.cos(azimuth_rad - width))
    lon1 = longitude + math.degrees(radius_lon * math.sin(azimuth_rad - width))

    # Second corner (main azimuth direction)
    lat2 = latitude + math.degrees(radius_lat * math.cos(azimuth_rad))
    lon2 = longitude + math.degrees(radius_lon * math.sin(azimuth_rad))

    # Third corner of the triangle
    lat3 = latitude + math.degrees(radius_lat * math.cos(azimuth_rad + width))
    lon3 = longitude + math.degrees(radius_lon * math.sin(azimuth_rad + width))

    return [[latitude, longitude], [lat1, lon1], [lat2, lon2], [lat3, lon3], [latitude, longitude]]
