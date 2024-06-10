#!/usr/bin/env python
# coding: utf-8

import os
import pandas as pd
import requests
import numpy as np
import json
import login
from datetime import datetime, timedelta
from openpyxl import load_workbook
import os
from geopy.geocoders import Nominatim
# ------------------ #
# Customer DB Params #
# ------------------ #

def get_cooler_df(company_id, user_name, company_df_path="company_Adriano.xlsx"):
    # Ler o DataFrame da empresa
    company_df = pd.read_excel(company_df_path)
    
    # Obter os detalhes da empresa com base no ID da empresa
    company = company_df[company_df["companyId"] == company_id]

    # Obter os parâmetros necessários
    db_url = company["endPoint"].values[0].replace('/parse/', '')
    app_id = company["appId"].values[0]

    # Realizar login
    session_token = login.login(user_name, db_url, app_id)

    header = {
        "Content-Type": "application/json",
        "X-Parse-Application-Id": app_id,
        'X-Parse-Session-Token': session_token
    }

    # Consulta dos coolers
    table = "/parse/classes/Cooler"
    field_list = [
        "coolerId",
        "usageStatus",
        "customPatrimonio",
        "controllerId",
        "oemSerial",
    ]
    url_params = {
        "keys": ",".join(field_list),
        "limit": "1000000"
    }

    response = requests.get(db_url + table, headers=header, params=url_params)
    cooler_data_json = response.json()["results"]
    cooler_df = pd.DataFrame.from_records(cooler_data_json, exclude=["createdAt", "updatedAt"])
    
    return cooler_df

company_id = 92
user_name = "adriano.alvesbitencourtdosanjos@gmail.com"
cooler_df = get_cooler_df(company_id, user_name)
print(cooler_df)