import os

import pandas as pd


def transform_htm_2_xls(htm_path: str, xls_path: str):
    html_data_list = pd.read_html(open(htm_path))
    table_data, table_header = list(), list()
    for html_data in html_data_list:
        table_header = list(list(html_data.values)[0])
        for row_data in list(html_data.values)[1:]:
            table_data.append(list(row_data))
    df = pd.DataFrame(table_data)
    df.columns = table_header
    ew = pd.ExcelWriter(xls_path)
    df.to_excel(ew, index=False)
    ew.save()


def transform_folder(file_path: str):
    [transform_htm_2_xls(os.path.join(file_path, file), os.path.join(file_path, file.replace('htm', 'xls'))) for file in
     os.listdir(file_path) if str.endswith(file, 'htm')]


if __name__ == '__main__':
    transform_folder(r'C:\Users\WWan95\Desktop')
