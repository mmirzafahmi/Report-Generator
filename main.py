import os
import glob
import argparse
import warnings

from datetime import date, timedelta
from Utils import PDF, DataOps

from pyunpack import Archive

warnings.simplefilter(action='ignore', category=FutureWarning)


if __name__ == '__main__':

    parser = argparse.ArgumentParser()

    parser.add_argument('input', type=str)
    parser.add_argument('output', type=str)

    args = parser.parse_args()

    if not os.path.exists(os.getcwd() + '\\logs'):
        os.makedirs(os.getcwd() + 'logs')
        os.makedirs(os.getcwd() + 'logs\\html')
        os.makedirs(os.getcwd() + 'logs\\img')
        os.makedirs(os.getcwd() + 'logs\\pdf')
        os.makedirs(os.getcwd() + 'logs\\spreadsheet')

    for dirs in os.listdir('logs'):

        DataOps.clean_dir(f'{os.getcwd()}\\logs\\{dirs}')

    log_dir = os.getcwd() + '\\logs'

    ss_dir = f'{log_dir}\\spreadsheet'
    pdf_dir = f'{log_dir}\\pdf'

    sources = glob.glob(f'{args.input}\\*.rar')

    for source in sources[2:]:

        print(source)
        Archive(source).extractall(ss_dir)
        sub_folder = os.listdir(ss_dir)[0] if len(os.listdir(ss_dir)) == 1 else ''
        spreadsheet_dir = f'{ss_dir}\\{sub_folder}' if sub_folder != '' else ss_dir
        e_files = glob.glob(f'{spreadsheet_dir}\\*.xlsx')
        release_date = (date.today() - timedelta(days=4)).strftime('%Y-%m-%d')
        spreadsheets = {
            'excel_files': [],
            'sheet_name': []
        }

        for f in e_files:

            DataOps.excel2pdf(
                path=f,
                date=release_date,
                log_dir=log_dir,
                pdf_dir=pdf_dir
            )

        f_name = source.split("\\")[-1][:-4]
        pdfs = glob.glob(pdf_dir + '\\*.pdf')

        for pdf in pdfs:
            PDF.Add_Watermark(
                input_file=pdf,
                output_file=pdf,
                watermark_file=f'{os.getcwd()}\\watermark.pdf'
            )

        PDF.Concat(
            input_files=pdfs,
            output=f'{args.output}\\{f_name}_{release_date}.pdf'
        )

        title_page = {'App': 2,
                      'CS': 3,
                      'PS': 4,
                      'Roaming In': 0,
                      'Roaming Out': 1,
                      'Roaming In Weekly': 5,
                      'Roaming Out Weekly': 6,
                      'CS Weekly': 7}

        PDF.Add_Title_Page(
            input_file=f'{args.output}\\{f_name}_{release_date}.pdf',
            title_file=f'{os.getcwd()}\\watermark.pdf',
            output_file=f'{args.output}\\{f_name}_{release_date}.pdf',
            page=title_page[f_name]
        )

        for dirs in os.listdir('logs'):
            DataOps.clean_dir(f'{os.getcwd()}\\logs\\{dirs}')
