from pathlib import Path

import xlwings as xw

import pandas as pd


class Wnioski:

    def __init__(self):
        # określenie lokalizacji pliku z wnioskami
        wnioski_dir = (Path(__file__) / ".." / ".." /
                       "src" / "outputs" / "wnioski").resolve()

        # określenie lokalizacji pliku z datami kursowania
        daty_kursowania_dir = (Path(__file__) / ".." / ".." /
                               "src" / "outputs" / "wnioski_daty_kursowania").resolve()

        # wczytanie wniosków z pliku excela do DateFrame
        for path in (wnioski_dir).rglob("*.xls*"):
            self.wb_wnioski = xw.Book(path)
            self.ws_wnioski = self.wb_wnioski.sheets[0]
            self.wnioski = self.ws_wnioski.range('A1').options(
                pd.DataFrame, expand='table').value

        # wczytanie dat kursowania z pliku excela do DateFrame
        for path in (daty_kursowania_dir).rglob("*.xls*"):
            self.wb_daty_kursowania = xw.Book(path)
            self.ws_daty_kursowania = self.wb_daty_kursowania.sheets[0]
            self.daty_kursowania = self.ws_daty_kursowania.range('A1').options(
                pd.DataFrame, expand='table').value

        self.wb_wnioski.app.quit()

    def all(self):

        return self.wnioski

    def all_daty_kursowania(self):

        return self.daty_kursowania

    def filtruj(self, wg_kolumny, wartosc):

        wnioski_wartosc = self.wnioski.loc[self.wnioski[wg_kolumny] == wartosc]

        return wnioski_wartosc

    def pobierz_do_dodatku(self, wg_kolumny, wartosc, nr_obiegu, rel_w_pot):

        df_daty_kursowania = pd.DataFrame()

        # pobierz wnioski do dodatku
        wnioski_wartosc = self.wnioski.loc[self.wnioski[wg_kolumny] == wartosc]

        # pobierz i posortuj daty kursowania pociągu
        for nr_zam in wnioski_wartosc["Nr zam."]:
            df_daty_kursowania = df_daty_kursowania.append(
                self.daty_kursowania.loc[nr_zam])

        df_daty_kursowania['Data kursowania'] = pd.to_datetime(
            df_daty_kursowania['Data kursowania'])

        df_daty_kursowania = df_daty_kursowania.reset_index()

        try:
            # wybierz z dat kursowania pierwszą datę dla każdego zamówienia
            df_daty_kursowania = df_daty_kursowania.drop_duplicates(subset=[
                                                                    'Nr zam.'])
        except:
            df_daty_kursowania.rename(
                columns={'index': 'Nr zam.'}, inplace=True)
            df_daty_kursowania = df_daty_kursowania.drop_duplicates(subset=[
                                                                    'Nr zam.'])

        # dodaj do wniosków, dla każdego zamówienia pierwszy dzień kursowania

        wnioski_wartosc = wnioski_wartosc.merge(
            df_daty_kursowania, how="inner", on=["Nr zam."])
        wnioski_wartosc = wnioski_wartosc.sort_values(by="Data kursowania")

        wnioski_wartosc = wnioski_wartosc.reset_index()

        # wybierz z wniosków tylko unikatowe wartości
        wnioski_unikatowe = wnioski_wartosc.drop_duplicates(
            subset=['Odj. RT', 'Prz. RT'])

        wnioski_unikatowe = wnioski_unikatowe.loc[:, ['Nr gr. poc.', 'Nr zam.', 'Odległość', 'Rodz. poc.',
                                                      'Nr poc.', 'Rel. od', 'Odj. RT', 'Rel. do', 'Prz. RT', 'Data kursowania']]

        wnioski_unikatowe = wnioski_unikatowe.reset_index(drop=True)

        wnioski_unikatowe.iloc[1:, 2] = 0

        # przypisz do wnioski nr obiegu
        wnioski_unikatowe['nr_obiegu'] = nr_obiegu

        # przypisz do wnioski nr relacji w planie obiegu taboru
        wnioski_unikatowe['rel_w_pot'] = rel_w_pot

        return wnioski_unikatowe

    def pobierz_liste_zamowien(self, nr_gr_poc, nr_poc):

        if nr_gr_poc == 0:
            wnioski_po_nr = self.wnioski.loc[self.wnioski["Nr poc."] == nr_poc]

        else:
            wnioski_po_nr = self.wnioski.loc[self.wnioski["Nr gr. poc."] == nr_gr_poc]

            if nr_poc != 0:
                mask = wnioski_po_nr['Nr poc.'] == nr_poc
                wnioski_po_nr = wnioski_po_nr.loc[mask, :]

        # wybierz z wniosków tylko unikatowe wartości
        wnioski_unikatowe = wnioski_po_nr.drop_duplicates(
            subset=['Nr zam.'])

        wnioski_unikatowe = wnioski_unikatowe.loc[:, [
            'Nr gr. poc.', 'Nr zam.']]

        return wnioski_unikatowe

    def filtruj_daty_kursowania(self, data):

        zamowienia = self.daty_kursowania.loc[self.daty_kursowania["Data kursowania"] == data]
        zamowienia = zamowienia.reset_index()

        return zamowienia
