import pandas as pd

from datetime import datetime

from openpyxl.utils.cell import get_column_letter

from scripts.isztp_date_conventer import zakres_dat


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

    def rysuj_kalendarz(self, daty, df_dates, xw_sheet, start_row, start_column):

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
