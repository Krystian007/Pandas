# -*- coding: utf-8 -*-

import pandas as pd
import datetime


def nowy_numer(string):

    string = str(string)
    n = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
    st = ''
    for _ in string:
        if _ in n:
            st += _

    return st


print('start: ', datetime.datetime.now())
plik_obce = pd.read_excel(
    '\\\\111.111.0.00\\users\\ksiegowosc_temp\\Makra\\Zamowienia.xlsx')
print('utworzenie pd: ', datetime.datetime.now())

plik_obce.opis = [nowy_numer(poz) for poz in plik_obce.opis]
print('dodanie nowych numerów: ', datetime.datetime.now())

plik_obce['numer_nowy'] = [poz for poz in plik_obce.opis]
print('dodanie nowych numerów na potrzeby makr: ', datetime.datetime.now())

plik_obce.to_excel('\\\\111.111.0.00\\users\\ksiegowosc_temp\\Makra\\Obce.xlsx',
                   header=False, index=False,)


print('koniec: ', datetime.datetime.now())
