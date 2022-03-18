# -*- coding: utf-8 -*-
"""
Created on Wed Oct 27 10:06:17 2021
processing of the payment file so that
they can be downloaded to the system
@author: kklos
"""
import datetime as dt
import pandas as pd
import numpy as np


print('start: ', dt.datetime.now())


def szukanie_numeru(szukana):
    """
    changing the payment number to the system order number

    Parameters
    ----------
    szukana : payment number.

    Returns
    -------
    system order number.

    """

    if szukana in dane.values:
        pozycja = np.argwhere(dane.values == szukana)
        return dane.values[pozycja[0, 0]]
    return szukana

def dzielenie(obiekt, nazwa):
    """
    creating a new payment file, ready to be downloaded to the system

    Parameters
    ----------
    obiekt : pd.read.csv
    nazwa : name of future file.

    """

    zbior = set(obiekt.Trans_ID)

    for paczka in zbior:
        nowy_plik = obiekt[obiekt.Trans_ID == paczka]
        nowy_plik.columns = [col.replace('_', ' ')
                             for col in nowy_plik.columns]
        dzien = nowy_plik['Data'].value_counts(normalize=True)
        dzien = dzien.index[0]
        kwota = round(nowy_plik.Kwota.sum(), 2)
        nowy_plik.Kwota = [str(poz).replace(',', '.')
                           for poz in nowy_plik.Kwota]
        nowy_plik.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\Pay Pro\\do '+
                           f'zaciągnięcia\\{dzien}_{paczka}_{nazwa}_{kwota}_PLN.xlsx',
                           index=False)


sprzedaz = pd.read_csv('C:\\Users\\kklos\\Desktop\\Programy\\Pay Pro\\pliki\\t.csv',
                       sep=',', usecols=(0, 2, 4, 6, 10, 11, 12))

sprzedaz = sprzedaz.rename(columns={'Sprzedawca': 'Typ_operacji', 'Przyjęcie': 'Data',
                                    'ID sesji': 'ID_sesji', 'Numer wypłaty': 'Trans_ID',
                                    'Klient': 'Imię_i_nazwisko'})

sprzedaz = sprzedaz.sort_index()

print('zaciągnięcie sprzedaży: ', dt.datetime.now())

zwroty = pd.read_csv('C:\\Users\\kklos\\Desktop\\Programy\\Pay Pro\\pliki\\z.csv',
                     sep=',', usecols=(1, 3, 4, 5, 8, 10))


zwroty = zwroty.rename(columns={'Sprzedawca': 'Typ_operacji', 'Data wykonania': 'Data',
                                'Tytuł': 'Opis', 'ID sesji': 'ID_sesji',
                                'ID wypłaty': 'Trans_ID'})
zwroty.Kwota = [-zl for zl in zwroty.Kwota]

print('zaciągnięcie zwrotów: ', dt.datetime.now())

zwroty['Imię_i_nazwisko'] = None

transakcje = sprzedaz.append(zwroty)


transakcje.Kwota = [float(zl/100) for zl in transakcje.Kwota]

dane = pd.read_excel('C:\\Users\\kklos\\Desktop\\dane.xlsx', usecols=(2, 3),
                     header=None, dtype=(str))

transakcje['Prowizja'] = None
transakcje['Saldo'] = None
transakcje['Order_ID'] = None
transakcje.Data = [str(dzien)[0:10] for dzien in transakcje.Data]
transakcje = transakcje[['Typ_operacji', 'Data', 'Trans_ID', 'Order_ID', 'Kwota', 'Prowizja',
                         'Saldo', 'Opis', 'Imię_i_nazwisko', 'ID_sesji']]


print('łączenie wszystkich transakcji: ', dt.datetime.now())

Militaria = transakcje[transakcje.Typ_operacji == 134748]


Militaria.Order_ID = Militaria.ID_sesji


dzielenie(Militaria, 'Militaria')

print('Militaria gotowe: ', dt.datetime.now())


Militaria_Shop = transakcje[transakcje.Typ_operacji == 134751]

order_id = []

for numer in Militaria_Shop.Opis:

    numer_ms = szukanie_numeru(numer)

    try:
        order_id.append(numer_ms[1])
    except ValueError:
        order_id.append(numer)


Militaria_Shop.Order_ID = order_id

dzielenie(Militaria_Shop, 'Militaria Shop')

print('Militaria Shop gotowe: ', dt.datetime.now())


Militaria_2 = transakcje[transakcje.Typ_operacji == 79847]


order_id = []

for numer in Militaria_2.Opis:

    try:
        order_id.append(numer[0:numer.index(' / ')])
    except ValueError:
        order_id.append(numer)

Militaria_2.Order_ID = order_id

dzielenie(Militaria_2, 'Militaria x')


print('Militaria 2 gotowe: ', dt.datetime.now())


print('Koniec: ', dt.datetime.now())
