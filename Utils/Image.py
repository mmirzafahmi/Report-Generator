import re
import matplotlib
import matplotlib.pyplot as plt

import win32com.client as win32

from PIL import ImageGrab


def Extract(excel_path, sheet_name, image_path, image_name):

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(excel_path, None, True)
    ws = workbook.Worksheets
    for sheet in ws:
        if sheet.Name == sheet_name:
            while True:
                try:
                    for i, shape in enumerate(sheet.Shapes):
                        if shape.Name.startswith('Picture'):
                            shape.Copy()
                            image = ImageGrab.grabclipboard()
                            image.save(f'{image_path}\\{image_name}.png', 'PNG')

                except AttributeError:
                    continue
                else:
                    break

    workbook = None
    excel = None


def Graph(df, columns, filename, save_dir):

    index, cols, vals = (None, None, None)
    for i in range(len(columns)):

        if re.search(r'\d{4}-\d{2}-\d{2}', str(df[columns[i]].values[0])) is not None:

            index = columns[i]

        elif str(df[columns[i]].values[0]).isalpha():

            cols = columns[i]

        else:

            vals = columns[i]

    if cols is None:
        cols = 'Service Application'

    ax = df.pivot(index=index,
                  columns=cols,
                  values=vals).plot(figsize=(12, 5), marker='.')

    ax.legend(
        ncol=len(df[cols].unique()),
        loc="lower left",
        bbox_to_anchor=(0, 1),
        fontsize='small',
        prop={'size': 6.9}
    )
    ax.get_yaxis().set_major_formatter(
        matplotlib.ticker.FuncFormatter(lambda x, p: format(int(x))))
    ax.set_ylabel('Number', loc='top')
    plt.xlabel('')
    plt.savefig(f'{save_dir}\\{filename}', dpi=100, bbox_inches='tight')
