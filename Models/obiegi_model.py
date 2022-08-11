from pathlib import Path

import xlwings as xw

import pandas as pd


class Obiegi:

    def __init__(self):
        # określenie lokalizacji pliku z wnioskami
        obiegi_dir = (Path(__file__) / ".." / ".." /
                      "src" / "inputs" / "obiegi").resolve()

        # wczytanie wniosków z pliku excela do DateFrame
        for path in (obiegi_dir).rglob("*.xls*"):
            self.wb_obiegi = xw.Book(path)
            self.ws_obiegi = self.wb_obiegi.sheets[0]
            self.obiegi = self.ws_obiegi.range('A1').options(
                pd.DataFrame, expand='table').value

        self.wb_obiegi.app.quit()

    def all(self):

        return self.obiegi

    def filtruj(self, wg_kolumny, wartosc):

        obiegi_wyfiltrowane = pd.DataFrame()

        if type(wartosc) == list:
            for element in wartosc:
                obiegi_wyfiltrowane = obiegi_wyfiltrowane.append(
                    self.obiegi.loc[self.obiegi[wg_kolumny] == element])

        else:

            obiegi_wyfiltrowane = self.obiegi.loc[self.obiegi[wg_kolumny] == wartosc]

        return obiegi_wyfiltrowane
