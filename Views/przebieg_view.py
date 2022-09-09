from email import header
import xlwings as xw

from xlwings.utils import rgb_to_int

import pandas as pd

import datetime as dt

from datetime import datetime

from env.env import FunkcjeGlobalne as fg


class PrzebiegView():

    def przebieg_do_xl(self, df_przebieg):

        print("Tworzenie pliku Excel...")

        wb_xl_przebieg = xw.Book()

        wb_xl_dodatek = xw.Book()

        nr_obiegow = df_przebieg.drop_duplicates(subset=['nr_obiegu'])

        nr_obiegow = nr_obiegow.loc[:, 'nr_obiegu']

        # określ ramy czasowe przebiegu
        p_start = datetime.strptime(
            df_przebieg.loc[:, 'Data'].min(), '%Y-%m-%d')

        p_end = datetime.strptime(
            df_przebieg.loc[:, 'Data'].max(), '%Y-%m-%d')

        p_zakres = pd.date_range(p_start, p_end)

        for nr_obiegu in range(int(nr_obiegow.min()), int(nr_obiegow.max() + 1)):

            df_przebieg_dla_obiegu = self.rozpisz_przebieg_obiegu(
                df_przebieg, nr_obiegu, p_zakres, p_end)

            df_przebieg_dla_obiegu = df_przebieg_dla_obiegu.sort_values(
                by=['nr_pojazdu', 'Data', 'Odj. RT'])

            self.rysuj_przebieg_do_xl(
                wb_xl_przebieg, df_przebieg_dla_obiegu, nr_obiegu)

            self.dodatek_z_przebiegu(
                wb_xl_dodatek, df_przebieg_dla_obiegu, nr_obiegu)

        print("Zakończono tworzenie pliku excel z sukcesem.")

    def dodatek_z_przebiegu(self, xw_book, df_przebieg_dla_obiegu, nr_obiegu):

        df = df_przebieg_dla_obiegu.copy()

        df_dates = df.drop_duplicates(subset="Data")

        df_dates = df_dates.set_index('Data')

        df_temp = pd.DataFrame()

        lista_df = []

        wariant = 1

        daty_wariantu = {}

        wystapienia = 0

        for index, row in df_dates.iterrows():

            # dla każdej daty w zakresie:

            df_c = df.copy()
            mask = (df_c['Data'] == index)
            df_temp = df_c.loc[mask, :]
            df_temp = df_temp.reset_index(drop=True)

            # rozdziel obiegi jeżeli są więcej niż jednodniowe:
            ilosc_pojazdow = df_przebieg_dla_obiegu.loc[:, 'dzien_w_obiegu'].max(
            )

            # jednodniowe obiegi dla każdego pojazdu
            for pojazd in range(1, int(ilosc_pojazdow) + 1):

                mask_pojazd = (df_temp['nr_pojazdu'] == pojazd)
                df_temp_p = df_temp.loc[mask_pojazd, :]

                df_temp_s_col = df_temp_p.loc[:, [
                    "Rel. od", "Odj. RT", "Rel. do", "Prz. RT"]]

                df_temp_s_col = df_temp_s_col.reset_index(drop=True)

                if len(lista_df) == 0:
                    lista_df.append(df_temp_p)
                    wystapienia = 1

                    daty_wariantu[0] = [
                        df_temp_p.iloc[0, 0]]

                else:
                    wystapienia = 0
                    for index_df_unique, df_unique in enumerate(lista_df):
                        df_s_col = df_unique.loc[:, ["Rel. od",
                                                     "Odj. RT", "Rel. do", "Prz. RT"]]

                        if df_temp_s_col.equals(df_s_col):
                            wystapienia = 1
                            daty_wariantu[index_df_unique].append(
                                df_temp_p.iloc[0, 0])

                if wystapienia == 0:
                    df_temp_p = df_temp_p.reset_index(drop=True)
                    try:
                        daty_wariantu[len(lista_df)] = [
                            df_temp_p.iloc[0, 0]]
                    except:
                        print(df_temp)
                        print(df_temp_p)
                        print(pojazd)

                    lista_df.append(df_temp_p)

                    # --------------------------------------------------------------------

        try:

            xw_book.sheets.add(
                f"obieg_{nr_obiegu}", after=f"obieg_{nr_obiegu - 1}")

        except:

            xw_book.sheets.add(f"obieg_{nr_obiegu}")

        f_row = 1

        for i, df_u in enumerate(lista_df):

            wariant = df_u.iloc[0, 0]

            df_u = df_u.loc[:, ["Odległość", "Nr poc.", "Rodz. poc.", "Rel. od",
                                "Odj. RT", "Rel. do", "Prz. RT", "Pojazdy"]]

            df_u.loc[:, 'wariant_obiegu'] = i

            ws_xl_dodatek_z_przebiegu = xw_book.sheets[f"obieg_{nr_obiegu}"]

            # WKLEJANIE DO EXCELA

            if i == 0:
                ws_xl_dodatek_z_przebiegu["A1"].options(
                    pd.DataFrame, expand='table', index=False).value = df_u

            else:

                f_row = f_row + l_row
                ws_xl_dodatek_z_przebiegu[f"A{f_row}"].expand('down').options(
                    pd.DataFrame, expand='table', index=False, header=False).value = df_u

            kalendarz = fg()
            kalendarz.rysuj_kalendarz(
                daty_wariantu[i], df_dates, ws_xl_dodatek_z_przebiegu, f_row, 11)

            df_u_len = len(df_u.index)

            if df_u_len < 14:
                l_row = 14
            else:
                l_row = df_u_len + 1

        # styl arkusza:
        ws_xl_dodatek_z_przebiegu["E:E"].number_format = 'gg:mm'
        ws_xl_dodatek_z_przebiegu["G:G"].number_format = 'gg:mm'
        ws_xl_dodatek_z_przebiegu.autofit(axis="columns")

    def rysuj_przebieg_do_xl(self, xw_book, df_przebieg_dla_obiegu, nr_obiegu):
        try:

            xw_book.sheets.add(
                f"obieg_{nr_obiegu}", after=f"obieg_{nr_obiegu - 1}")

        except:

            xw_book.sheets.add(f"obieg_{nr_obiegu}")

        ws_xl_przebieg = xw_book.sheets[f"obieg_{nr_obiegu}"]

        ws_xl_przebieg["A1"].options(
            pd.DataFrame, expand='table', index=False).value = df_przebieg_dla_obiegu

        ws_xl_przebieg["I1"].expand("down").number_format = 'gg:mm'
        ws_xl_przebieg["K1"].expand("down").number_format = 'gg:mm'
        ws_xl_przebieg.autofit(axis="columns")

        # sprawdzenie czy w przebiegu pociąg kończy i zaczyna w tych samych stacjach
        for row in range(ws_xl_przebieg["A2"].expand("down").last_cell.row, 3, -1):
            st_o = ws_xl_przebieg[f"H{row}"].value
            st_p = ws_xl_przebieg[f"J{row - 1}"].value

            pojazd_o = ws_xl_przebieg[f"M{row}"].value
            pojazd_p = ws_xl_przebieg[f"M{row - 1}"].value

            if pojazd_o != pojazd_p:
                continue

            if st_o != st_p:
                print(f"{st_o} != {st_p}")
                ws_xl_przebieg.range(f"H{row}").color = (255, 0, 0)
                ws_xl_przebieg.range(f"J{row - 1}").color = (255, 0, 0)

    def rozpisz_przebieg_obiegu(self, df_przebieg, nr_obiegu, p_zakres, p_end):

        print(f"rysuje przebieg obiegu : {nr_obiegu}")

        df_przebieg_dla_obiegu = df_przebieg.loc[df_przebieg['nr_obiegu'] == nr_obiegu]

        df_przebieg_dla_obiegu = df_przebieg_dla_obiegu.sort_values(
            by=['Data', 'Odj. RT'])

        # jeżeli dzien w obiegu nie jest zdefiniowany to oznacz jako 1
        dfd = df_przebieg_dla_obiegu.copy()
        mask = dfd.dzien_w_obiegu.isnull()
        dfd.loc[mask, 'dzien_w_obiegu'] = 1

        df_przebieg_dla_obiegu = dfd

        ilosc_pojazdow = df_przebieg_dla_obiegu.loc[:, 'dzien_w_obiegu'].max(
        )

        if ilosc_pojazdow == 1:
            df_przebieg_dla_obiegu.insert(12, "nr_pojazdu", 1)
        else:
            df_przebieg_dla_obiegu['nr_pojazdu'] = pd.Series()
            df_przebieg_dla_obiegu.insert(13, "wykorzystanie", 0)
            # rozdziel obieg pomiędzy pojazdy, jeżeli obieg jest kilkudniowy
            for pojazd in range(1, int(ilosc_pojazdow) + 1):
                nr_dnia_obiegu = pojazd
                for p_dzien in p_zakres:

                    dfd = df_przebieg_dla_obiegu.copy()
                    mask = (dfd['Data'] == datetime.strftime(
                        p_dzien, '%Y-%m-%d')) & (dfd['dzien_w_obiegu'] == nr_dnia_obiegu)
                    dfd.loc[mask, 'nr_pojazdu'] = pojazd
                    dfd.loc[mask, 'wykorzystanie'] = 1
                    df_przebieg_dla_obiegu = dfd

                    # 1. nast_dzien_o następny dzień obiegu
                    nast_dzien_o = nr_dnia_obiegu + 1

                    if nast_dzien_o > ilosc_pojazdow:
                        nast_dzien_o = 1

                    if not p_dzien == p_end:

                        # sprawdź czy następnego dnia zaczyna w stacji, w której skończył
                        while not self.przejscie_nocne_pojazdu(
                                df_przebieg_dla_obiegu, p_dzien, nast_dzien_o, nr_dnia_obiegu):

                            if nast_dzien_o == nr_dnia_obiegu:
                                if self.przejscie_nocne_pojazdu(
                                        df_przebieg_dla_obiegu, p_dzien, nast_dzien_o, nr_dnia_obiegu) == False:
                                    print(
                                        f"BŁĄD_PRZEJŚCIA: Nie poprawne przejście jednoski z dnia {p_dzien} na dzień następny!")
                                    break

                            nast_dzien_o = nast_dzien_o + 1
                            if nast_dzien_o > ilosc_pojazdow:
                                nast_dzien_o = 1

                    nr_dnia_obiegu = nast_dzien_o

        return df_przebieg_dla_obiegu

    def przejscie_nocne_pojazdu(self, df_przebieg_dla_obiegu, p_dzien, nr_dnia_obiegu, poprz_nr_dnia_ob):

        dfd = df_przebieg_dla_obiegu.copy()
        mask = (dfd['Data'] == datetime.strftime(
            p_dzien, '%Y-%m-%d')) & (dfd['dzien_w_obiegu'] == poprz_nr_dnia_ob)
        dfe = dfd.loc[mask, :]
        ostatnia_stacja = dfe.iloc[-1, 9]
        o_s_nr_poc = dfe.iloc[-1, 5]

        dff = df_przebieg_dla_obiegu.copy()
        mask = (dff['Data'] == datetime.strftime(
            (p_dzien + dt.timedelta(days=1)), '%Y-%m-%d')) & (dff['dzien_w_obiegu'] == nr_dnia_obiegu)
        dfg = dff.loc[mask, :]
        pierwsza_nast_doba_stacja = dfg.iloc[0, 7]
        wykorzystanie = dfg.iloc[0, 13]

        if ostatnia_stacja != pierwsza_nast_doba_stacja:
            return False
        else:
            if wykorzystanie == 1:
                return False
            return True
