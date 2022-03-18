# -*- coding: utf-8 -*-
"""
Created on Wed Oct 27 10:06:17 2021
a simple program to generate a column from the content that
facilitates the identification of orders for the needs of vba excel
@author: kklos
"""
import datetime as dt
import pandas as pd


def nowy_numer(string):
    """
    Separate only intiger value from stirng

    Parameters
    ----------
    string : string

    Returns
    -------
    nowy_string : string of intiger.

    """

    string = str(string)
    numer = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
    nowy_string = ''
    for _ in string:
        if _ in numer:
            nowy_string += _

    return nowy_string


print('start: ', dt.datetime.now())
plik_obce = pd.read_excel(
    '\\\\111.111.0.00\\users\\ksiegowosc_temp\\Makra\\Zamowienia.xlsx')
print('utworzenie pd: ', dt.datetime.now())

plik_obce.opis = [nowy_numer(poz) for poz in plik_obce.opis]
print('dodanie nowych numerów: ', dt.datetime.now())

plik_obce['numer_nowy'] = plik_obce.opis
print('dodanie nowych numerów na potrzeby makr: ', dt.datetime.now())

plik_obce.to_excel('\\\\111.111.0.00\\users\\ksiegowosc_temp\\Makra\\Obce.xlsx',
                   header=False, index=False,)


print('koniec: ', dt.datetime.now())
