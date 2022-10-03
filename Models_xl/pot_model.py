from pathlib import Path

import xlwings as xw

import pandas as pd


class Pot:

    def __init__(self):
        # określenie lokalizacji pliku z wnioskami
        pot_dir = (Path(__file__) / ".." / ".." /
                   "src" / "inputs" / "plan_obiegow_taboru").resolve()

        # wczytanie wniosków z pliku excela do DateFrame
        for path in (pot_dir).rglob("*.xls*"):
            self.wb_pot = xw.Book(path)
            self.ws_pot = self.wb_pot.sheets[0]
            self.pot = self.ws_pot.range('A1').options(
                pd.DataFrame, expand='table').value

            #self.pot = self.pot.reset_index()

        self.wb_pot.app.quit()

    def all(self):

        return self.pot

    def filtruj(self, wg_kolumny, wartosc):

        pot_wyfiltrowane = pd.DataFrame()

        if type(wartosc) == list:
            for element in wartosc:
                pot_wyfiltrowane = pot_wyfiltrowane.append(
                    self.pot.loc[self.pot[wg_kolumny] == element])

        else:

            pot_wyfiltrowane = self.pot.loc[self.pot[wg_kolumny] == wartosc]

        df_temp = pot_wyfiltrowane.fillna(0)

        pot_wyfiltrowane = df_temp

        return pot_wyfiltrowane
