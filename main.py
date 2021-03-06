from xml.etree.ElementTree import Element, SubElement, Comment, tostring
import xml.etree.cElementTree as ET
from io import StringIO
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
import pandas as pd
from math import log10
import requests, base64
from flask import escape
from flask import Flask, render_template, Response
from base64_encoder import *
from bs4 import BeautifulSoup
import io, os

def quickbase_auth():
    url = 'https://soteriarf.quickbase.com/db/main?a=API_authenticate&username=evan@thecbrgroup.com&password=AnUEh7so&hours=24'
    r = requests.get(url)
    soup = BeautifulSoup(r.content,'xml')
    for url in soup.find_all('ticket'):
            auth_ticket = url.text
    return auth_ticket

def W2dBm(W):
    return 10*log10(W*1000)
def dBm2W(dBm):
    return 10**((dBm)/10) / 1000

def upload_file_to_quickbase(table_id, record_id, fid, out_file):
    # Create XML payload
    root = ET.Element("qdbapi")
    udata = ET.SubElement(root, 'udata')
    udata.text = 'mydata'
    ticket = ET.SubElement(root, 'ticket')
    ticket.text = quickbase_auth()
    apptoken = ET.SubElement(root, 'apptoken')
    apptoken.text = 'dz2abv5chpk6bs35syzkcydqjfb'
    ET.SubElement(root, "field", fid=str(fid), filename=out_file).text = encode_file(out_file)
    rid = ET.SubElement(root, 'rid')
    rid.text = str(record_id)
    tree = ET.ElementTree(root)
    out = io.BytesIO()
    tree.write(out)
    xml_str = out.getvalue().decode()
    # API_UploadFile
    url = "https://soteriarf.quickbase.com/db/" + table_id
    payload = xml_str
    headers = {
      'Content-Type': 'application/xml',
      'QUICKBASE-ACTION': 'API_UploadFile'
    }
    response = requests.request("POST", url, headers=headers, data = payload)
    print(response.text.encode('utf8'))

def qb_dataframe(table_id):
    url = "https://api.quickbase.com/v1/fields?tableId=" + table_id
    payload = {}
    headers = {
      'QB-Realm-Hostname': 'soteriarf',
      'Authorization': 'QB-USER-TOKEN b4pp4v_paqb_eyxwg9bec3v8pb63hjwjb9t36ii',
      'Content': 'application.json'
    }
    response = requests.request("GET", url, headers=headers, data = payload)
    df = pd.json_normalize(response.json())
    df = df.sort_values('id').reset_index(drop=True)
    df = df[df['fieldType'] != 'user']
    df = df[df['fieldType'] != 'file']
    label_list = list(df['label'])
    id_list = list(df['id'])
    body = {"from": table_id, "select": id_list, "options": {"skip": 0, "top": 10000}}
    r = requests.post('https://api.quickbase.com/v1/records/query', headers = headers, json = body)
    df = pd.json_normalize(r.json()['data'])
    df.columns = df.columns.str.replace(".value", "")
    id_list = [str(i) for i in id_list] 
    for col in df.columns:
        df.rename(columns={col:label_list[id_list.index((col))]}, inplace=True)
    return df

def antenna_inventory(request):
    request_args = request.args
    site_id = request_args['site_id']
    record_id = request_args['record_id']
    address = request_args['address']

    # Load Transmitters Table, query on site_id
    df = qb_dataframe('bqzaae4x9')
    df = df[df['Antenna Location - SITE_ID'] == site_id]
    
    df = df.rename(columns={
        'Related Antenna Location': 'antenna_id',
        'Antenna Location - SECTOR': 'sector',
        'Port': 'port',
        'Technology': 'technology',
        'Frequency MHz': 'freq',
        'Pattern Code - Manufacturer': 'manufacturer',
        'Antenna Location - ANTENNA_MODEL': 'model',
        'Antenna Location - AZIMUTH': 'azimuth',
        'Antenna Location - RAD_CENTER_FT': 'agl',
        'Input Power (W)': 'power',
        'Pattern Code - Gain (dBd)': 'gain'
    })

    df = df[['antenna_id', 'sector', 'port', 'technology', 'freq', 'manufacturer', 'model', 'azimuth', 'agl', 'power', 'gain']]

    df['erp_dbm'] = ''
    df['erp'] = ''
    for x in range(len(df)):
        df['erp_dbm'].iloc[x] = round(W2dBm(df['power'].iloc[x]*1000) + df['gain'].iloc[x], 1)
        df['erp'].iloc[x] = int(dBm2W(df['erp_dbm'].iloc[x]) / 1000)
    df = df.drop('erp_dbm', 1)
    df['freq'] = df['freq'].astype(int)
    df['azimuth'] = df['azimuth'].astype(int)
    df['power'] = df['power'].astype(int)
    df = df.round({'gain': 1})
    df.sort_values(['antenna_id', 'port', 'freq'], ascending=[True, True, True], inplace=True)

    row_contents = []
    for x in range(len(df)):
        row_contents.append(eval(df.iloc[x].to_json()))

    context = {
        'site_id': site_id,
        'address': address,
        'row_contents': row_contents
    }

    out_file = '/tmp/AntennaInventory.docx'
    doc = DocxTemplate('AntennaInventoryTpl.docx')
    doc.render(context)
    doc.save(out_file)

    tmp_files = os.listdir('/tmp')
    

    upload_file_to_quickbase('bqq5m3crs', record_id, 249, out_file)

    return ('AntennaInventory.docx created. Reload Quick Base to access the file.')