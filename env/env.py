import pandas as pd

from datetime import datetime

from openpyxl.utils.cell import get_column_letter

import re


class ZmienneSrodowiskowe:

    def __init__(self):

        self.swieta_poza_niedziela = [
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

    def swieta_poza_niedziela(self):
        return self.swieta_poza_niedziela


class FunkcjeGlobalne:

    def __init__(self):
        self.zmienneSrodowiskowe = ZmienneSrodowiskowe()

    def rysuj_kalendarz(self, daty, df_dates, xw_sheet, start_row, start_column):

        df_dates = df_dates.sort_index()

        # definicja zakresu kolumn
        start_column_0 = get_column_letter(start_column)
        start_column_l = get_column_letter(start_column + 1)
        end_column_l = get_column_letter(start_column + 38)

        xw_sheet.range(f"{start_column_l}:{end_column_l}").column_width = 2
        xw_sheet.range(
            f"{start_column_l}:{end_column_l}").font.size = 7

        # Nagłówek z dni tygodnia i zakoloruj weekendy:
        dni_tygodnia = ['pn', 'wt', 'sr', 'cz', 'pt', 'sb', 'nd']
        i = 0
        for j in range(start_column + 1, start_column + 38):
            xw_sheet.range(
                f"{get_column_letter(j)}{start_row}").value = dni_tygodnia[i]

            # Koloruj sobotę i niedzielę
            if i == 5:
                xw_sheet.range(
                    f"{get_column_letter(j)}{start_row}:{get_column_letter(j)}{start_row + 13}").color = (250, 250, 250)
            if i == 6:
                xw_sheet.range(
                    f"{get_column_letter(j)}{start_row}:{get_column_letter(j)}{start_row + 13}").color = (245, 245, 245)

            i += 1
            if i == 7:
                i = 0

        start_date = datetime.strptime(
            df_dates.first_valid_index(), '%Y-%m-%d')

        start_date_m = int(start_date.strftime('%m'))

        start_date_y = int(start_date.strftime('%Y'))

        start_date = datetime.strptime(
            f"{start_date_y}-{start_date_m}-01", '%Y-%m-%d')

        end_date = datetime.strptime(df_dates.last_valid_index(), '%Y-%m-%d')

        zakres_dat = pd.date_range(start_date, end_date)

        miesiac = start_date.strftime('%m')

        wiersz = start_row + 1

        kolumna = start_column + 1

        xw_sheet.range(f"{start_column_0}{wiersz}").value = miesiac

        ilosc_w_kalend = 1

        for dzien in zakres_dat:

            if miesiac != dzien.strftime('%m'):
                wiersz += 1
                kolumna = start_column + 1
                xw_sheet.range(
                    f"{start_column_0}{wiersz}").value = dzien.strftime('%m')
                miesiac = dzien.strftime('%m')
                ilosc_w_kalend += 1

            if int(dzien.strftime('%d')) == 1:
                kolumna = kolumna + dzien.weekday()

            xw_sheet.range(
                f"{get_column_letter(kolumna)}{wiersz}").value = dzien.strftime('%d')

            if dzien.strftime('%Y-%m-%d') in daty:
                xw_sheet.range(
                    f"{get_column_letter(kolumna)}{wiersz}").color = (0, 255, 0)

            kolumna += 1

        # ramki
        whole_calendar_range = xw_sheet.range(
            f"{get_column_letter(start_column + 1)}{start_row + 1}:{get_column_letter(start_column + 38)}{start_row + ilosc_w_kalend}")
        whole_calendar_range.api.Borders(7).Weight = 3
        whole_calendar_range.api.Borders(8).Weight = 3
        whole_calendar_range.api.Borders(9).Weight = 3
        whole_calendar_range.api.Borders(10).Weight = 3
        whole_calendar_range.api.Borders(11).Weight = 1
        whole_calendar_range.api.Borders(12).Weight = 1

    def kursuje_w_obiegu(self, termin, dzien):

        # Sprawdź czy pociąg w tym dniu kursuje w tym obiegu (po zadanym terminie)
        if termin == "H":
            return True

        elif termin == "D":
            # w dni robocze
            if dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela:
                return False
            elif (dzien.isoweekday() == 6) | (dzien.isoweekday() == 7):
                return False
            else:
                return True

        elif termin == "C":
            # w weekendy i święta
            if dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela:
                return True
            elif (dzien.isoweekday() == 6) | (dzien.isoweekday() == 7):
                return True
            else:
                return False

        elif termin == "A":
            # od poniedziałku do piątku
            if (dzien.isoweekday() == 6) | (dzien.isoweekday() == 7):
                return False
            else:
                return True

        elif termin == "B":
            # bez soboty
            if (dzien.isoweekday() == 6):
                return False
            else:
                return True

        elif termin == "E":
            # od poniedziałku do soboty oprócz świąt
            if dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela:
                return False
            elif (dzien.isoweekday() == 7):
                return False
            else:
                return True

        elif re.search(r"^\[\d-\d\]\-$", termin):
            # dni tygodnia od - do
            od_do = re.findall(r"\d", termin)
            if (dzien.isoweekday() >= int(od_do[0])) & (dzien.isoweekday() <= int(od_do[1])) & (dzien not in self.zmienneSrodowiskowe.swieta_poza_niedziela):
                return True
            else:
                return False

        elif re.search(r"^\[\d-\d\]$", termin):
            # dni tygodnia od - do
            od_do = re.findall(r"\d", termin)
            if (dzien.isoweekday() >= int(od_do[0])) & (dzien.isoweekday() <= int(od_do[1])):
                return True
            else:
                return False

        elif re.search(r"^\[\d\]$", termin):
            # jeden dzien w tygodniu
            dzien_w_tyg = re.findall(r"\d", termin)
            if dzien.isoweekday() == int(dzien_w_tyg[0]):
                return True
            else:
                return False

        elif re.search(r"^\[\d\]\+$", termin):
            # jeden dzien w tygodniu
            dzien_w_tyg = re.findall(r"\d", termin)
            if (dzien.isoweekday() == int(dzien_w_tyg[0])) | (dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela):
                return True
            else:
                return False

        # specjalne wyjątki w terminie kursowania:-------------------------------------------------------------------

        elif re.search(r"^\[\d\]\*3$", termin):
            # jeden dzien w tygodniu
            dzien_w_tyg = re.findall(r"\d", termin)
            spec_dzien = [datetime(2022, 12, 27)]

            if (dzien.isoweekday() == int(dzien_w_tyg[0])) | (dzien in spec_dzien):
                return True
            else:
                return False

        elif re.search(r"^\[\d\]\*4$", termin):
            # jeden dzien w tygodniu
            dzien_w_tyg = re.findall(r"\d", termin)
            spec_dzien = [datetime(2023, 1, 5)]
            spec_dzien_2 = [datetime(2023, 1, 6)]

            if ((dzien.isoweekday() == int(dzien_w_tyg[0])) | (dzien in spec_dzien)) & (dzien not in spec_dzien_2):
                return True
            else:
                return False

        # elif re.search(r"^\[\d-\d\]\*2$", termin):
        #     # dni tygodnia od - do
        #     spec_dzien = [datetime(2022, 11, 11)]
        #     od_do = re.findall(r"\d", termin)
        #     if (dzien.isoweekday() >= int(od_do[0])) & (dzien.isoweekday() <= int(od_do[1])) & (dzien not in spec_dzien):
        #         return True
        #     else:
        #         return False

        elif re.search(r"^H\*1$", termin):
            spec_dzien = [datetime(2022, 12, 25)]

            if dzien in spec_dzien:
                return False
            else:
                return True

        elif re.search(r"^\*1$", termin):
            spec_dzien = [datetime(2022, 12, 25)]

            if dzien in spec_dzien:
                return True
            else:
                return False

        elif re.search(r"^H\*2$", termin):
            spec_dzien = [datetime(2022, 12, 25)]

            if dzien in spec_dzien:
                return False
            else:
                return True

        elif re.search(r"^\*2$", termin):
            spec_dzien = [datetime(2022, 12, 25)]

            if dzien in spec_dzien:
                return True
            else:
                return False

        # -----------------------------------------------------------------------------------------------------------

        else:
            print(
                f"brak rozważanego terminu kursowania {termin}")
            return False
