import os

import pandas as pd

price_list_dict = {
    'CN': '50                  China',
    'HK': '50                  China',
    'TW': '59                  Taiwan'
}

dt_type_dict = {
    4: float,
    6: float
}


def format_dt_type(df):
    for k, v in dt_type_dict.items():
        df[k] = df[k].head(5).append(df[k].tail(-5).astype(v))


def transform_htm_2_xls(htm_path: str, xls_path: str, price_list_name: str):
    html_data_list = pd.read_html(open(htm_path))
    table_data = [
        ['Price List', None, None, price_list_dict.get(price_list_name)],
        [], [],
        [None, 'CnTy', 'Material', None, ' Amount', 'Unit', 'per', 'UoM', 'Valid From', 'Valid to', 'D'],
        []
    ]
    for html_data in html_data_list:
        for row_data in list(html_data.values)[1:]:
            table_data.append([None, *list(row_data)[:2], None, *list(row_data)[2:]])
    df = pd.DataFrame(table_data)
    format_dt_type(df)
    ew = pd.ExcelWriter(xls_path)
    df.to_excel(ew, index=False, header=False)
    ew.save()
    os.remove(htm_path)
