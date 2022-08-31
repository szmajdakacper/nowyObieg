"""
isztp_date_conventer
Skrypt, który zamienia pseudoczytelne daty generowane przez ISZTP
na listę dat w formacie yyyy-mm-dd
"""

# IMPORTOWANIE MODUŁÓW: ---------------------------------------------------------------------------

from pathlib import Path

import xlwings as xw

import pandas as pd

import re

from datetime import datetime

# 1. DEKLARACJA ZMIENNYCH I STAŁYCH ---------------------------------------------------------------


miesiace_rzymskie = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5,
                     "VI": 6, "VII": 7, "VIII": 8, "IX": 9,
                     "X": 10, "XI": 11, "XII": 12}

rok_rrj = 2022

swieta_poza_niedziela = [
    datetime(2022, 1, 6),
    datetime(2022, 4, 18),
    datetime(2022, 5, 3),
    datetime(2022, 6, 16),
    datetime(2022, 8, 15),
    datetime(2022, 11, 1),
    datetime(2022, 11, 11),
    datetime(2022, 12, 26),
    datetime(2023, 1, 6),
    datetime(2023, 4, 10),
    datetime(2023, 5, 1),
    datetime(2023, 5, 3),
    datetime(2023, 6, 8),
    datetime(2023, 8, 15),
    datetime(2023, 11, 1),
]

# 2. DEFINICJA FUNKCJI: ---------------------------------------------------------------------------


# 1) funkcja obsługująca: Zakres dat

def zakres_dat(data_kursowania):

    daty_zakresu = []

    lista_caly_zakres = []

    lista_dat = []

    oraz = []
    oprocz = []

    data_kursowania_podzielona = data_kursowania.split(" ")

    for fragment_daty in data_kursowania_podzielona:

        fragment_daty = fragment_daty.strip()

        if fragment_daty == '':
            continue
        else:

            # najpierw pobieramy dwie daty zakresu od - do
            if re.search(r"^\d{1,2}\.", fragment_daty):

                fragment_daty = fragment_daty.split("-")

                for data in fragment_daty:

                    data = data.strip()

                    if data == '':
                        continue
                    else:
                        daty_zakresu.append(pojedyncza_data(data))

                # Tworzymy listę dni w zakresie od - do
                start_date = daty_zakresu[0]
                end_date = daty_zakresu[1]

                lista_caly_zakres = pd.date_range(start_date, end_date)

    # po zdefiniowaniu zakresu, usuwamy daty i zostawiamy wykluczenia (jeżeli są)
    del data_kursowania_podzielona[0]
    def_wykluczen = " ".join(data_kursowania_podzielona)

    if def_wykluczen == "w (D)":
        oraz = [1, 2, 3, 4, 5]
        for dzien_kursowania in lista_caly_zakres:
            if (dzien_kursowania.isoweekday() in oraz) & (dzien_kursowania not in swieta_poza_niedziela):
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "w (A)":
        oraz = [1, 2, 3, 4, 5]
        for dzien_kursowania in lista_caly_zakres:
            if dzien_kursowania.isoweekday() in oraz:
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "w (C)":
        oraz = [6, 7]
        for dzien_kursowania in lista_caly_zakres:
            if (dzien_kursowania in swieta_poza_niedziela) | (dzien_kursowania in oraz):
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "w (B)":
        oprocz = [6]
        for dzien_kursowania in lista_caly_zakres:
            if dzien_kursowania.isoweekday() not in oprocz:
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "w (E)":
        oprocz = [7]
        for dzien_kursowania in lista_caly_zakres:
            if (dzien_kursowania.isoweekday() not in oprocz) & (dzien_kursowania not in swieta_poza_niedziela):
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "codziennie oprócz świąt":
        for dzien_kursowania in lista_caly_zakres:
            if dzien_kursowania not in swieta_poza_niedziela:
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "w niedziele i święta":
        for dzien_kursowania in lista_caly_zakres:
            oraz = [7]
            if (dzien_kursowania in swieta_poza_niedziela) | (dzien_kursowania in oraz):
                lista_dat.append(dzien_kursowania)

    elif re.search(r"^(\(\d\))+$", def_wykluczen):
        oraz = re.findall(r"\d", def_wykluczen)
        for i, v in enumerate(oraz):
            oraz[i] = int(v)

        for dzien_kursowania in lista_caly_zakres:
            if dzien_kursowania.isoweekday() in oraz:
                lista_dat.append(dzien_kursowania)

    elif re.search(r"oprócz świąt$", def_wykluczen):
        oraz = re.findall(r"\d", def_wykluczen)
        for i, v in enumerate(oraz):
            oraz[i] = int(v)

        for dzien_kursowania in lista_caly_zakres:
            if (dzien_kursowania.isoweekday() in oraz) & (dzien_kursowania not in swieta_poza_niedziela):
                lista_dat.append(dzien_kursowania)

    elif re.search(r"oraz w święta$", def_wykluczen):
        oraz = re.findall(r"\d", def_wykluczen)
        for i, v in enumerate(oraz):
            oraz[i] = int(v)

        for dzien_kursowania in lista_caly_zakres:
            if (dzien_kursowania.isoweekday() in oraz) & (dzien_kursowania in swieta_poza_niedziela):
                lista_dat.append(dzien_kursowania)

    elif re.search(r"^codziennie oprócz", def_wykluczen):
        oprocz = re.findall(r"\d", def_wykluczen)
        for i, v in enumerate(oprocz):
            oprocz[i] = int(v)

        for dzien_kursowania in lista_caly_zakres:
            if dzien_kursowania.isoweekday() not in oprocz:
                lista_dat.append(dzien_kursowania)

    elif def_wykluczen == "":
        lista_dat = lista_caly_zakres

    else:
        print(def_wykluczen)

    return lista_dat


# 2) funkcja obsługująca: dwie daty

def dwie_daty(data_kursowania):

    lista_dat = list()

    daty_podzielone = data_kursowania.split(",")

    for data in daty_podzielone:

        data = data.strip()

        if data == '':
            continue

        else:
            lista_dat.append(pojedyncza_data(data))

    return lista_dat


# 3) funkcja obsługująca: pojedyńcza data

def pojedyncza_data(data_kursowania):
    dzien = re.findall("^\d{1,2}", data_kursowania)
    dzien = int(dzien[0])

    miesiac = re.findall("\.\w{1,3}", data_kursowania)
    miesiac = miesiac[0][1:]
    miesiac = miesiace_rzymskie[miesiac]

    # sprawdzam po ilości kropek "." w stringu
    # czy data jest zapisana w formacie: 26.XII.22 czy 10.IV
    ilosc_kropek = re.findall("\.", data_kursowania)

    if len(ilosc_kropek) > 1:
        rok = re.findall("\d{2,4}$", data_kursowania)
        rok = int("20" + rok[0])

    else:
        rok = rok_rrj

    return datetime(int(rok), int(miesiac), int(dzien))


# 3. SKRYPT: --------------------------------------------------------------------------------------


# określenie lokalizacji pliku z wnioskami
wnioski_dir = (Path(__file__) / ".." / ".." /
               "src" / "outputs" / "wnioski").resolve()

# wczytanie wniosków z pliku excela
for path in (wnioski_dir).rglob("*.xls*"):
    wb_wnioski = xw.Book(path)
    sheet = wb_wnioski.sheets['Sheet1']
    wnioski = sheet.range('A1').options(
        pd.DataFrame, expand='table').value

# unikatowe indeksowanie wniosków
wnioski = wnioski.reset_index(drop=True)
wnioski.index.names = ["Lp."]

# Dodaj kolumnę kursuje_lista_dat
wnioski["Kursuje_lista_dat"] = ""


# Stwórz dodatkową relacyjną tabele, która będzie przechowywała "Nr zam." i "Data Kursowania", dla każdego dnia osobno
tabela_relacyjna_kursowania = []

# Główna pętla po wszystkich wnioskach
for index, row in wnioski.iterrows():

    lista_dat = list()

    # terminy:
    # usuwam z wartości w kolumnie 'Kursuje' tekst zawarty w nawiasach kwadratowych
    terminy = re.sub(r'\[.*\]', '', row['Kursuje'])

    # podział terminów na daty i zakresy
    # każdy termin kursowania może składać się
    # z pojedyńczych dni, lub zakresów, rozdzielonych ";".
    daty_kursowania = terminy.split(";")

    # pętla po każdej, wydzielonej, dacie kursowania
    for data_kursowania in daty_kursowania:

        data_kursowania = data_kursowania.strip()

        if data_kursowania == '':
            continue

        elif re.search(r"^\d{1,2}\.", data_kursowania):

            if re.search(r"-", data_kursowania):
                # 1) data_kursowania zawiera zakres dat, od '-' do'
                # zakres_dat(data_kursowania)
                for jedna_data in zakres_dat(data_kursowania):
                    lista_dat.append(jedna_data)
                    tabela_relacyjna_kursowania.append(
                        [row['Nr zam.'], jedna_data.strftime("%Y.%m.%d")])

            elif re.search(r",", data_kursowania):
                # 2) data_kursowania zawiera dwie daty oddzielone przecinkiem ','
                # dwie_daty(data_kursowania)
                for jedna_data in dwie_daty(data_kursowania):
                    lista_dat.append(jedna_data)
                    tabela_relacyjna_kursowania.append(
                        [row['Nr zam.'], jedna_data.strftime("%Y.%m.%d")])

            else:
                # 3) kursuje tylko w jeden dzień
                lista_dat.append(pojedyncza_data(data_kursowania))
                tabela_relacyjna_kursowania.append(
                    [row['Nr zam.'], pojedyncza_data(data_kursowania).strftime("%Y.%m.%d")])

        else:
            lista_dat.append("Nieznany format daty!!!")

    for element_listy_dat in lista_dat:
        if wnioski.at[index, 'Kursuje_lista_dat'] == "":
            wnioski.at[index, 'Kursuje_lista_dat'] = element_listy_dat.strftime(
                "%Y.%m.%d")
        else:
            wnioski.at[index, 'Kursuje_lista_dat'] = wnioski.at[index, 'Kursuje_lista_dat'] + \
                "," + element_listy_dat.strftime("%Y.%m.%d")


# Zapisz wnioski z listą dat kursowania.
sheet["A1"].options(pd.DataFrame, header=1, index=True,
                    expand='table').value = wnioski

sheet["A1"].expand("right").api.Font.Bold = True
sheet["A1"].expand("down").api.Font.Bold = True
sheet["A1"].expand("right").api.Borders.Weight = 2
sheet["A1"].expand("down").api.Borders.Weight = 2
sheet["A1"].expand().api.WrapText = False
sheet["A1"].expand().autofit()
sheet["A1"].expand().api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

fn_wb_wnioski = wb_wnioski.name

wb_wnioski.save(Path(__file__) / ".." / ".." /
                "src" / "outputs" / "wnioski" / fn_wb_wnioski)

wb_wnioski.app.quit()


# zapisz relacyjną tabele do pliku excela:
df_rel_tab_kursuje = pd.DataFrame(tabela_relacyjna_kursowania, columns=[
    "Nr zam.", "Data kursowania"])
wb_kursuje = xw.Book()
sheet_kursuje = wb_kursuje.sheets[0]
sheet_kursuje["A1"].options(pd.DataFrame, header=1, index=False,
                            expand='table').value = df_rel_tab_kursuje

fn_wb_kursuje = "kursuje" + fn_wb_wnioski

wb_kursuje.save(Path(__file__) / ".." / ".." /
                "src" / "outputs" / "wnioski_daty_kursowania" / fn_wb_kursuje)

wb_kursuje.app.quit()
