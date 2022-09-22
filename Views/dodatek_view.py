from pathlib import Path

import xlwings as xw

from xlwings.utils import rgb_to_int

import pandas as pd


class DodatekView():

    def dodatek_do_xl(self, df_dodatek):

        print("Tworzenie pliku Excel...")

        tab_widoku_df_dodatek = df_dodatek.loc[:, ['Nr gr. poc.', 'opis_obiegu', 'Odległość',
                                                   'Rodz. poc.', 'Nr poc.', 'Termin', 'Uwagi', 'Rel. od', 'Odj. RT', 'Rel. do', 'Prz. RT', 'Zestawienie']]

        wb_xl_dodatek = xw.Book()

        ws_xl_dodatek = wb_xl_dodatek.sheets[0]

        ws_xl_dodatek["A1"].options(
            pd.DataFrame, expand='table', index=False).value = tab_widoku_df_dodatek

        # Stylowanie dodatku:

        ws_xl_dodatek["I1"].expand("down").number_format = 'gg:mm'
        ws_xl_dodatek["K1"].expand("down").number_format = 'gg:mm'
        ws_xl_dodatek.autofit(axis="columns")
        ws_xl_dodatek["A1"].expand(
            "table").api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

        last_cell_row = ws_xl_dodatek["A1"].expand("down").last_cell.row + 2
        zakres_dodatku = ws_xl_dodatek.range(
            f"A1:L{last_cell_row}")

        # krawędzie wewnętrzne
        zakres_dodatku.api.Borders(11).Weight = 1

        zakres_dodatku.api.Borders(12).Weight = 1

        # Lewa krawędź
        zakres_dodatku.api.Borders(7).Weight = 3
        # Górna krawędź
        zakres_dodatku.api.Borders(8).Weight = 3
        # Dolna krawędź
        zakres_dodatku.api.Borders(9).Weight = 3
        ws_xl_dodatek["A1"].expand("table").api.Borders(9).Weight = 3
        # Prawa krawędź
        zakres_dodatku.api.Borders(10).Weight = 3

        ws_xl_dodatek.range(f"A1:A{last_cell_row}").api.Borders(10).Weight = 3

        ws_xl_dodatek["A1"].expand("right").api.Borders(9).Weight = 3
        ws_xl_dodatek["A1"].expand("right").api.Borders(10).Weight = 3
        ws_xl_dodatek["A1"].expand("right").api.Font.Bold = True

        # dwie ostatnie bez podziałki
        ws_xl_dodatek.range(
            f"A{last_cell_row - 1}:L{last_cell_row}").color = (255, 255, 255)
        ws_xl_dodatek.range(
            f"A{last_cell_row - 1}:A{last_cell_row}").api.Borders(12).LineStyle = 0
        ws_xl_dodatek.range(
            f"A{last_cell_row - 1}:A{last_cell_row}").api.Borders(11).LineStyle = 0
        ws_xl_dodatek.range(
            f"B{last_cell_row - 1}:L{last_cell_row}").api.Borders(12).LineStyle = 0
        ws_xl_dodatek.range(
            f"B{last_cell_row - 1}:L{last_cell_row}").api.Borders(11).LineStyle = 0

        # podkreśl warianty
        for odl in ws_xl_dodatek["C1"].expand("down"):
            if ws_xl_dodatek.range(f"C{odl.row}").value == 0:
                ws_xl_dodatek.range(
                    f"A{odl.row}:L{odl.row}").color = (248, 255, 229)
                ws_xl_dodatek.range(
                    f"A{odl.row}:L{odl.row}").api.Font.Color = rgb_to_int((102, 102, 102))

        # rozdziel obiegi
        kom_poczatkowa = ws_xl_dodatek["B2"].expand("down").last_cell.value
        kom_poprzednia = ws_xl_dodatek["B2"].expand("down").last_cell.value

        for row in range(ws_xl_dodatek["B2"].expand("down").last_cell.row, 2, -1):
            opis_obiegu = ws_xl_dodatek[f"B{row}"]
            if opis_obiegu.value == kom_poczatkowa:
                continue

            elif opis_obiegu.value == None:
                continue

            else:
                kom_poczatkowa = opis_obiegu.value
                kom_poprzednia = ws_xl_dodatek[f"B{row + 1}"]
                kom_poprzednia.api.EntireRow.Insert()
                ws_xl_dodatek.range(
                    f"A{kom_poprzednia.row - 1}:L{kom_poprzednia.row - 1}").color = (255, 255, 255)

                ws_xl_dodatek.range(
                    f"A{kom_poprzednia.row}:L{kom_poprzednia.row}").api.Borders(8).Weight = 3
                kom_poprzednia.api.EntireRow.Insert()
                ws_xl_dodatek.range(
                    f"A{opis_obiegu.row}:L{opis_obiegu.row}").api.Borders(9).Weight = 3

                ws_xl_dodatek.range(
                    f"B{kom_poprzednia.row - 1}:L{kom_poprzednia.row - 1}").api.Borders(11).LineStyle = 0
                ws_xl_dodatek.range(
                    f"B{kom_poprzednia.row}:L{kom_poprzednia.row}").api.Borders(11).LineStyle = 0

        ws_xl_dodatek["A1"].api.EntireRow.Insert()
        ws_xl_dodatek["A1"].api.EntireColumn.Insert()

        wb_xl_dodatek.save(Path(__file__) / ".." / ".." /
                           "src" / "outputs" / "pot" / "pot.xlsx")

        print("Proces tworzenia dodatku w Excelu zakończony sukcesem!")
