from io import StringIO
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
import pandas as pd
from math import log10
import requests

def W2dBm(W):
    return 10*log10(W*1000)
def dBm2W(dBm):
    return 10**((dBm)/10) / 1000

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
    
    return 'test123'