import xlwings as xw

from xlwings.utils import rgb_to_int

import pandas as pd

from datetime import datetime


class PrzebiegView():

    def przebieg_do_xl(self, df_przebieg):

        print("Tworzenie pliku Excel...")

        nr_obiegow = df_przebieg.drop_duplicates(subset=['nr_obiegu'])

        nr_obiegow = nr_obiegow.loc[:, 'nr_obiegu']

        wb_xl_przebieg = xw.Book()

        # określ ramy czasowe przebiegu
        p_start = datetime.strptime(
            df_przebieg.loc[:, 'Data'].min(), '%Y-%m-%d')

        p_end = datetime.strptime(
            df_przebieg.loc[:, 'Data'].max(), '%Y-%m-%d')

        p_zakres = pd.date_range(p_start, p_end)

        for nr_obiegu in range(int(nr_obiegow.min()), int(nr_obiegow.max() + 1)):

            print(f"rysuje przebieg obiegu : {nr_obiegu}")

            df_przebieg_dla_obiegu = df_przebieg.loc[df_przebieg['nr_obiegu'] == nr_obiegu]

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
                # rozdziel obieg pomiędzy pojazdy, jeżeli obieg jest kilkudniowy
                for pojazd in range(1, int(ilosc_pojazdow) + 1):
                    nr_dnia_obiegu = pojazd
                    for p_dzien in p_zakres:

                        dfd = df_przebieg_dla_obiegu.copy()
                        mask = (dfd['Data'] == datetime.strftime(
                            p_dzien, '%Y-%m-%d')) & (dfd['dzien_w_obiegu'] == nr_dnia_obiegu)
                        dfd.loc[mask, 'nr_pojazdu'] = pojazd
                        df_przebieg_dla_obiegu = dfd

                        # zapamiętaj poprzedni nr dnia obiegu
                        poprz_nr_dnia_ob = nr_dnia_obiegu

                        # zmień dzień obiegu dla pojazdu:
                        nr_dnia_obiegu = nr_dnia_obiegu + 1
                        if nr_dnia_obiegu > ilosc_pojazdow:
                            nr_dnia_obiegu = 1

                        # sprawdź czy następnego dnia zaczyna w stacji, w której skończył
                        while not self.przejscie_pojazdu(
                                df_przebieg_dla_obiegu, p_dzien, nr_dnia_obiegu, poprz_nr_dnia_ob):

                            if nr_dnia_obiegu == poprz_nr_dnia_ob:
                                if self.przejscie_pojazdu(
                                        df_przebieg_dla_obiegu, p_dzien, nr_dnia_obiegu, poprz_nr_dnia_ob) == False:
                                    print(
                                        f"Nie poprawne przejście jednoski z dnia {p_dzien} na dzień następny!")
                                    break

                            nr_dnia_obiegu = nr_dnia_obiegu + 1
                            if nr_dnia_obiegu > ilosc_pojazdow:
                                nr_dnia_obiegu = 1

            try:

                wb_xl_przebieg.sheets.add(
                    f"obieg_{nr_obiegu}", after=f"obieg_{nr_obiegu - 1}")

            except:

                wb_xl_przebieg.sheets.add(f"obieg_{nr_obiegu}")

            ws_xl_przebieg = wb_xl_przebieg.sheets[f"obieg_{nr_obiegu}"]

            ws_xl_przebieg["A1"].options(
                pd.DataFrame, expand='table', index=False).value = df_przebieg_dla_obiegu

            ws_xl_przebieg["I1"].expand("down").number_format = 'gg:mm'
            ws_xl_przebieg["K1"].expand("down").number_format = 'gg:mm'
            ws_xl_przebieg.autofit(axis="columns")

        print("Zakończono tworzenie pliku excel z sukcesem.")

    def przejscie_pojazdu(self, df_przebieg_dla_obiegu, p_dzien, nr_dnia_obiegu, poprz_nr_dnia_ob):

        dfd = df_przebieg_dla_obiegu.copy()
        mask = (dfd['Data'] == datetime.strftime(
            p_dzien, '%Y-%m-%d')) & (dfd['dzien_w_obiegu'] == poprz_nr_dnia_ob)
        dfe = dfd.loc[mask, :]
        ostatnia_stacja = dfe.iloc[-1, 9]
        o_s_nr_poc = dfe.iloc[-1, 5]
        dff = df_przebieg_dla_obiegu.copy()
        mask = (dff['Data'] == datetime.strftime(
            p_dzien, '%Y-%m-%d')) & (dff['dzien_w_obiegu'] == nr_dnia_obiegu)
        dfg = dff.loc[mask, :]
        pierwsza_nast_doba_stacja = dfg.iloc[0, 7]
        p_s_nr_poc = dfg.iloc[0, 5]

        if ostatnia_stacja != pierwsza_nast_doba_stacja:
            print(
                f"{datetime.strftime(p_dzien, '%d.%m')}. {o_s_nr_poc}: {ostatnia_stacja} nierówna się {p_s_nr_poc}: {pierwsza_nast_doba_stacja}")
            return False
        else:
            return True
