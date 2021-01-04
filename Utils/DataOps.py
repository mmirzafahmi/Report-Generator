import os
import pdfkit
import shutil
import dateutil

import numpy as np
import pandas as pd

from Utils import Image, HTML
from dateutil.parser import parse


def clean_dir(folder):

    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


def read_xlsx(xlsx):

    xl = pd.ExcelFile(xlsx)

    for sheet_name in xl.sheet_names:
        if not sheet_name.startswith('Conditions'):
            table = xl.parse(sheet_name)
            is_columns = [True if c.startswith('Unnamed') else False for c in table.columns]

            if any(is_columns):
                df = pd.DataFrame()
                for i, row in table.iterrows():
                    if row.notnull().all():
                        df = table.iloc[(i + 1):].reset_index(drop=True)
                        df.columns = list(table.iloc[i])
                        break

                return df, sheet_name
            else:
                return table, sheet_name


def get_pivot_params(df):

    def get_datetime_column(frame):

        for column in frame.columns:

            try:
                frame[column].apply(lambda s: parse(s))
                return column
            except (TypeError, ValueError, dateutil.parser.ParserError):
                continue

        return ''

    tmp = df.copy()
    tmp.replace('--', np.nan, inplace=True)

    params = dict()
    params['columns'] = ''
    params['index'] = get_datetime_column(df)
    params['values'] = '' if len(df.columns) == 3 else []

    for col in tmp.columns:

        if col != params['index']:

            try:

                tmp[col].apply(lambda v: float(v))

            except (TypeError, ValueError):

                params['columns'] = col
                pass

            finally:

                if type(params['values']) == str:

                    params['values'] = col if col not in params.values() else ''

                else:

                    params['values'].append(col if col not in params.values() else '')

    return params


def processed_dataframe(df):

    params = get_pivot_params(df)

    n_unique = df[params.get('columns')].nunique() if params.get('columns') != '' else 20

    conditions = [
        '' not in params.values(),
        not isinstance(params.get('values'), list),
        n_unique <= 15
    ]

    if all(conditions):

        try:
            df[params.get('values')] = pd.to_numeric(df[params.get('values')])
        except (KeyError, ValueError):
            pass

        pivot = df.pivot(
            index=params.get('index'),
            columns=params.get('columns'),
            values=params.get('values')
        ).rename_axis(None, axis=1).reset_index()

        return pivot

    else:

        return df


def excel2pdf(path, date, log_dir, pdf_dir):

    img_dir = f'{log_dir}\\img'
    html_dir = f'{log_dir}\\html'

    df, sheet_name = read_xlsx(path)
    processed_df = processed_dataframe(df)

    f_name = path.split('\\')[-1][:-5]
    Image.Extract(path, sheet_name=sheet_name, image_path=img_dir, image_name=f_name)

    html_table = HTML.build_table(processed_df, 'blue_light', font_size='7px')

    html = f"""
    <!DOCTYPE html> 
    <html>
    <head>
        <style>
            td {'white-space:nowrap'}
        </style>
        <title>Report Automation</title>
    </head>
    <body>
    <h3>Filename: {f_name}</h3>
    <h3>Sheetname: {sheet_name}</h3>
    <h3>Release date: {date}</h3>
    <img src="{img_dir}\\{f_name}.png", width="650" height="400">
    {html_table}
    </body>
    <html>
    """

    with open(f"{html_dir}\\{f_name}_{sheet_name}.html", "w") as Html_file:
        Html_file.write(html)

    path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf, )
    options = {
        "enable-local-file-access": None,
        "quiet": ''
    }

    pdfkit.from_file(
        f"{html_dir}\\{f_name}_{sheet_name}.html",
        f'{pdf_dir}\\{f_name}_{sheet_name}.pdf',
        configuration=config,
        options=options
    )
