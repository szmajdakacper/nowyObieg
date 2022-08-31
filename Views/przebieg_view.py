import xlwings as xw

from xlwings.utils import rgb_to_int

import pandas as pd


class PrzebiegView():

    def przebieg_do_xl(self, df_przebieg):

        print("Tworzenie pliku Excel...")

        nr_obiegow = df_przebieg.drop_duplicates(subset=['nr_obiegu'])

        nr_obiegow = nr_obiegow.loc[:, 'nr_obiegu']

        wb_xl_przebieg = xw.Book()

        for nr_obiegu in range(int(nr_obiegow.min()), int(nr_obiegow.max() + 1)):

            print(f"rysuje przebieg obiegu : {nr_obiegu}")

            df_przebieg_dla_obiegu = df_przebieg.loc[df_przebieg['nr_obiegu'] == nr_obiegu]

            try:

                wb_xl_przebieg.sheets.add(
                    f"obieg_{nr_obiegu}", after=f"obieg_{nr_obiegu - 1}")

            except:

                wb_xl_przebieg.sheets.add(f"obieg_{nr_obiegu}")

            ws_xl_przebieg = wb_xl_przebieg.sheets[f"obieg_{nr_obiegu}"]

            ws_xl_przebieg["A1"].options(
                pd.DataFrame, expand='table', index=False).value = df_przebieg_dla_obiegu

        print("Zakończono tworzenie pliku excel z sukcesem.")
