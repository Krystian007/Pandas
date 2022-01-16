# -*- coding: utf-8 -*-
"""
Created on Thu Nov 25 07:56:33 2021

@author: kklos
"""

import pandas as pd
import numpy as np
import datetime


def upperstr(poz):

    if type(poz) == str:
        return(poz.upper())


def nowy_numer(string):

    string = str(string)
    n = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
    st = ''
    for _ in string:
        if _ in n:
            st += _

    return st


def szukanie_249(szukana, baza):

    if szukana in baza.values:
        numer = np.argwhere(baza.values == szukana)
        return baza.values[numer[0, 0]]


def log(tekst):

    print(f'>>>>> {tekst}', datetime.datetime.now())


log('Start programu')
obce = pd.read_excel('\\\\111.111.1.11\\users\\ksiegowosc_temp\\Makra\\Obce.xlsx',
                     header=None, dtype=(str), usecols=(1, 2, 3, 4, 5, 6))
obce.columns = ['zamowienie', 'faktura',
                'kontrahent', 'konto', 'numer', 'liczby']
log('Utworzenie obcych')


plik_249 = pd.read_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\pliki\\249.xlsx',
                         dtype=(str))
plik_249.columns = ['data', 'dowod', 'konto',
                    'wn', 'ma', 'przeciwstawne', 'tresc']
plik_249.tresc = [upperstr(poz) for poz in plik_249.tresc]
plik_249.wn = [float(wn) for wn in plik_249.wn]
plik_249.ma = [float(ma) for ma in plik_249.ma]
plik_249['liczby'] = [nowy_numer(poz) for poz in plik_249.tresc]
log('Utworzenie pliku 249')

mtr = []
konto = []

for opis in plik_249.values:

    numery = pd.Series(szukanie_249(opis[6], obce))

    if numery.all() != []:
        mtr.append(numery[1])
        konto.append(f'200-{numery[3]}')

    else:

        if len(opis[7]) > 6:

            numery = pd.Series(szukanie_249(opis[7], obce))

            if numery.all() != []:
                mtr.append(numery[1])
                konto.append(f'200-{numery[3]}')
            else:
                mtr.append(np.nan)
                konto.append(np.nan)

        else:
            mtr.append(np.nan)
            konto.append(np.nan)


plik_249['faktura'] = mtr
plik_249['numer_konta'] = konto

zrobione = pd.notnull(plik_249['numer_konta'])
eksport = plik_249[zrobione]

niezrobione = pd.isnull(plik_249['numer_konta'])
plik_249 = plik_249[niezrobione]

pozycje_248_nierowne = []
pozycje_248 = []

podwojne = pd.Series(plik_249['liczby'].value_counts())
podwojne = [podwojne.index[pod]
            for pod in range(podwojne.size) if podwojne[pod] >= 2]


for pozycja in podwojne:

    podwojne_pozycje = plik_249['liczby'].isin([pozycja])
    spr = pd.DataFrame(plik_249[podwojne_pozycje])
    if spr['wn'].sum() == spr['ma'].sum():
        pozycje_248.append(pozycja)
    else:
        pozycje_248_nierowne.append(pozycja)


zakres_248 = plik_249['liczby'].isin([_ for _ in pozycje_248])
eksport_248 = plik_249[zakres_248]
eksport_248['faktura'] = '248'
eksport_248['numer_konta'] = '248'

plik_249 = plik_249[-zakres_248]

pozycje_248_nierowne.remove('')
zakres_248_nierowne = plik_249['liczby'].isin(
    [_ for _ in pozycje_248_nierowne])
eksport_248_nierowne = plik_249[zakres_248_nierowne]
eksport_248_nierowne['faktura'] = '248_nierowne'
eksport_248_nierowne['numer_konta'] = '248_nierowne'

plik_249 = plik_249[-zakres_248_nierowne]

log('Utworzenie plików eksportu 248')

obce_binarny = pd.read_excel('\\\\192.168.0.30\\users\\ksiegowosc_temp\\OBCE\\Obce_binarny.xlsb',
                             header=None, dtype=(str), usecols=(0, 1, 2), engine='pyxlsb')
obce_binarny.columns = ['zamowienie', 'numer', 'kontrahent']

log('Utworzenie obce binarne')

obce_binarny['liczby'] = [nowy_numer(poz) for poz in obce_binarny.numer]
log('dodanie liczb w obce binarny')

mba = []
kh = []

for opis in plik_249.values:

    numery = pd.Series(szukanie_249(opis[6], obce_binarny))

    if numery.all() != []:
        mba.append(numery[0])
        kh.append(numery[2])
        print('znaleziono')

    else:

        if len(opis[7]) > 6:

            numery = pd.Series(szukanie_249(opis[7], obce_binarny))

            if numery.all() != []:
                mba.append(numery[0])
                kh.append(numery[2])
                print('znaleziono')
            else:
                mba.append(opis[6])
                kh.append(np.nan)

        else:
            mba.append(opis[6])
            kh.append(np.nan)


plik_249['faktura'] = mba
plik_249['numer_konta'] = kh

pozycje_248_nierowne = []
pozycje_248 = []

podwojne = pd.Series(plik_249['faktura'].value_counts())
podwojne = [podwojne.index[pod]
            for pod in range(podwojne.size) if podwojne[pod] >= 2]


for pozycja in podwojne:

    podwojne_pozycje = plik_249['faktura'].isin([pozycja])
    spr = pd.DataFrame(plik_249[podwojne_pozycje])
    if spr['wn'].sum() == spr['ma'].sum():
        pozycje_248.append(pozycja)
    else:
        pozycje_248_nierowne.append(pozycja)


zakres_248 = plik_249['faktura'].isin([_ for _ in pozycje_248])
eksport_248_xxx = plik_249[zakres_248]
eksport_248_xxx['numer_konta'] = '248'

plik_249 = plik_249[-zakres_248]

try:
    pozycje_248_nierowne.remove('')
except:
    pass

zakres_248_nierowne = plik_249['faktura'].isin(
    [_ for _ in pozycje_248_nierowne])
eksport_248_nierowne_xxx = plik_249[zakres_248_nierowne]
eksport_248_nierowne_xxx['numer_konta'] = '248_nierowne'

plik_249 = plik_249[-zakres_248_nierowne]

log('Utworzenie plików eksportu 249 xxx')

eksport.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\gotowe pliki\\200.xlsx',
                 header=False, index=False,)
eksport_248.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\gotowe pliki\\248.xlsx',
                     header=False, index=False,)
eksport_248_nierowne.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\gotowe pliki\\248_nierowne.xlsx',
                              header=False, index=False,)
eksport_248_xxx.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\gotowe pliki\\248_xxx.xlsx',
                         header=False, index=False,)
eksport_248_nierowne_xxx.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\gotowe pliki\\248_nierowne_xxx.xlsx',
                                  header=False, index=False,)
plik_249.to_excel('C:\\Users\\kklos\\Desktop\\Programy\\249\\gotowe pliki\\249.xlsx',
                  header=False, index=False,)


log('Eksport zakończony')

log('Koniec programu')
