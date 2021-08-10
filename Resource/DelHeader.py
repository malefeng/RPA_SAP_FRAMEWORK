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
    html_data_list = pd.read_html(htm_path)
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


def merge_html_2_xls(file_path: str, file_name_str: str, scope_name: str,final_file_name: str):
    table_data = [
        ['Price List', None, None, price_list_dict.get(scope_name)],
        [], [],
        [None, 'CnTy', 'Material', None, ' Amount', 'Unit', 'per', 'UoM', 'Valid From', 'Valid to', 'D'],
        []
    ]
    file_name_list = file_name_str.split(',')
    if len(file_name_list) == 0:
        return
    for file_name in file_name_list:
        html_data_list = pd.read_html(os.path.join(file_path, file_name))
        for html_data in html_data_list:
            for row_data in list(html_data.values)[1:]:
                table_data.append([None, *list(row_data)[:2], None, *list(row_data)[2:]])
    df = pd.DataFrame(table_data)
    format_dt_type(df)
    ew = pd.ExcelWriter(os.path.join(file_path, final_file_name))
    df.to_excel(ew, index=False, header=False)
    ew.save()
    for file_name in file_name_str.split(','):
        os.remove(os.path.join(file_path, file_name))


if __name__ == '__main__':
    file_path = r'C:\Users\WWan95\Desktop\SAP merge\test'
    file_names = '224247_HO22 Product Pricing_CN.htm,224317_HO22 Product Pricing_CN.htm,224347_HO22 Product Pricing_CN.htm,224418_HO22 Product Pricing_CN.htm,224449_HO22 Product Pricing_CN.htm'
    scope_name = 'CN'
    final_file_name = 'HO22 Product Pricing_CN.xls'
    merge_html_2_xls(file_path,file_names,scope_name,final_file_name)