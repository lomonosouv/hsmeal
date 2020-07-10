import json
import os
from string import ascii_lowercase

import numpy as np
from openpyxl import load_workbook


def get_nrows(worksheet):
    ws1 = worksheet
    nrows = len(ws1['D'])
    for i in range(len(ws1['D'])):
        if ws1['D'][nrows - i - 1].value is not None:
            nrows = nrows - i
            break

    return nrows


def get_names_columns():
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(cur_dir, '../../data/subscribers.json'), 'r') as f:
        subscribers = json.load(f)

    names = [sub["excel_title"] for sub in subscribers]

    alphabet = list(ascii_lowercase)
    start = alphabet.index('g')
    end = start + len(names)

    return names, alphabet[start: end]


def get_prices(excel_name):
    wb = load_workbook(excel_name, data_only=True)
    ws1 = wb.active

    nrows = get_nrows(ws1)
    names, cols = get_names_columns()

    total_price = ws1['f'][nrows].value

    res_names = []
    res_prices = []

    for name, col in zip(names, cols):
        value = float(ws1[col][nrows].value)

        if value > 0:
            res_names.append(name)
            res_prices.append(value)

    res_names = np.array(res_names)
    res_prices = np.array(res_prices)

    assert sum(res_prices) == total_price

    if total_price < 1000:
        res_prices += np.ceil(150 / len(res_prices))

    elif total_price < 1500:
        res_prices += np.ceil(100 / len(res_prices))

    return dict(zip(res_names, res_prices))

