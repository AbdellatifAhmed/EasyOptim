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

output_dir = os.path.join(os.getcwd(), 'OutputFiles')
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
    dmp_data_LnCel_DF = Lnrel_audit_form['LnCel']
    dmp_data_LnCel_DF.columns = dmp_data_LnCel_DF.columns.str.strip()
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

    lnrel_Performance_DF['Target Cell'] = lnrel_Performance_DF['eci_id'].apply(get_target_cell_id)
    lnrel_Performance_DF['Target eNB'] = lnrel_Performance_DF['eci_id'].apply(get_target_enb_id)
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

    out_columns = ['MRBTS','LNBTS','LNCEL','LNREL','Target eNB','Target Cell','Distance','cellIndOffNeigh','handoverAllowed','removeAllowed','Inter eNB HO attempts','Inter eNB HO Fails','Intra eNB HO attempts','Intra eNB HO Fails']
    output_Table = lnrel_Performance_DF[out_columns]
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

def clean_Sites_db(sites_df):
    sites_df.columns = sites_df.columns.str.strip()
    sites_df['NodeB Id'] = sites_df['NodeB Name'].apply(lambda x: x.split('-')[0])
    sites_df["polygon"] = sites_df.apply(lambda row: calculate_sector(row["Lat"], row["Long"], row["Bore"]), axis=1)
    sites_df["coordinates"] = sites_df.apply(lambda row: get_coordinate(row["Long"], row["Lat"]), axis=1)
    return sites_df

def calculate_sector(latitude, longitude, azimuth, radius=0.001):
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

    return [[longitude, latitude], [lon1, lat1], [lon2, lat2], [lon3, lat3], [longitude, latitude]]

def get_coordinate(longitude,latitude):
    return [longitude, latitude]