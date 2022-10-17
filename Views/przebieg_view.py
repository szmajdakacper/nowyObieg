from pathlib import Path

import xlwings as xw

from xlwings.utils import rgb_to_int

import pandas as pd

import re

from env.env import FunkcjeGlobalne as fg


class PrzebiegView():

    def __init__(self):
        self.df_do_wykresow = pd.DataFrame(data=[], columns=["nr_wykresu", "Odległość", "Rodz. poc.", "Nr poc.", "Termin", "Uwagi", "Rel. od",
                                                             "Odj. RT", "Rel. do", "Prz. RT", "Pojazdy", "opis_obiegu", 'wariant_obiegu'])

    def pokaz_przebieg_df(self, przebieg_df):

        wb_xl_przebieg = xw.Book()
        wb_xl_pot_roz = xw.Book()

        wb_temp_p = xw.Book()
        ws_temp = wb_temp_p.sheets[0]

        ws_temp["A1"].expand('down').options(
            pd.DataFrame, expand='table', index=False).value = przebieg_df

        # najmniejszy i największy nr obiegu
        p_nr_obiegu = int(przebieg_df['nr_obiegu'].min())
        o_nr_obiegu = int(przebieg_df['nr_obiegu'].max())

        for nr_obiegu in range(o_nr_obiegu, p_nr_obiegu - 1, -1):

            print(
                f"Rysowanie przebiegów: Postęp {round(((o_nr_obiegu - nr_obiegu) / o_nr_obiegu) * 100, 1)} %")

            maska_obiegu = przebieg_df['nr_obiegu'] == nr_obiegu
            df_obieg = przebieg_df.loc[maska_obiegu, :]

            opis_obiegu = df_obieg.iloc[0, 3]

            # -------------------------------
            # Stwórz dodatek (pot) rozszerzony
            df_obieg_pot_r = df_obieg.copy()
            df_obieg_pot_r['nr_pojazdu'] = 1
            self.dodatek_z_przebiegu(
                wb_xl_pot_roz, df_obieg_pot_r, nr_obiegu, opis_obiegu)
            # -------------------------------

            df_obieg = self.rozpisz_pojazdy(df_obieg)

            # stwórz arkusz dla obiegu i zapisz przebieg do xl
            wb_xl_przebieg.sheets.add(f"obieg_{nr_obiegu}")
            ws_xl_przebieg_obiegu = wb_xl_przebieg.sheets[f"obieg_{nr_obiegu}"]

            ws_xl_przebieg_obiegu["A1"].expand('down').options(
                pd.DataFrame, expand='table', index=False).value = df_obieg

            # formatowanie arkusza
            self.rysuj_przebieg_do_xl(wb_xl_przebieg, df_obieg, nr_obiegu)

        # Zapisz do plków Excela

        wb_xl_przebieg.save(Path(__file__) / ".." / ".." /
                            "src" / "outputs" / "pot" / "przebieg.xlsx")

        wb_xl_pot_roz.save(Path(__file__) / ".." / ".." /
                           "src" / "outputs" / "pot" / "pot_rozszerzony.xlsx")

        # Makro do rysowania wykresów obiegu:
        self.wklej_df_do_wykr()

    def rozpisz_pojazdy(self, df_obieg):

        df_obieg_c = df_obieg.copy()

        df_obieg_c = df_obieg_c.sort_values(
            by=['Data', 'dzien_w_obiegu', 'Odj. RT'])

        ilosc_pojazdow = int(df_obieg_c.loc[:, 'dzien_w_obiegu'].max())

        # jeżeli obieg jest jednodniowy to nie trzeba go rozpisywać
        if ilosc_pojazdow <= 1:
            df_obieg_c['dzien_w_obiegu'] = 1
            df_obieg_c['nr_pojazdu'] = 1
            return df_obieg_c

        df_obieg_c['wykorzystano'] = 0

        for pojazd in range(1, ilosc_pojazdow + 1):

            stacja_postoju = ""
            postoj_od_godz = ""
            dzien_postoju = ""
            post_w_obiegu = ""

            pomin_dzien_obiegu = 0

            for i, pociag in df_obieg_c.iterrows():

                # Pomiń jeżeli pociąg jest w innym obiegu
                if pociag['wykorzystano'] == 1:
                    continue

                # Rozpoczęcie przebiegu dla tego pojazdu
                if stacja_postoju == "":
                    stacja_postoju = pociag['Rel. do']
                    postoj_od_godz = pociag['Prz. RT']
                    postoj_od_godz = self.sprawdz_format_godz(postoj_od_godz)
                    dzien_postoju = pociag['Data']
                    post_w_obiegu = pociag['dzien_w_obiegu']

                    df_obieg_c.loc[i, 'nr_pojazdu'] = pojazd
                    df_obieg_c.loc[i, 'wykorzystano'] = 1
                    continue

                # Przypisz cechy pociągu do zmiennych

                stacja_odjazdu = pociag['Rel. od']

                godzina_odjazdu = pociag['Odj. RT']
                godzina_odjazdu = self.sprawdz_format_godz(godzina_odjazdu)

                dzien_odjazdu = pociag['Data']

                dzien_w_obiegu = pociag['dzien_w_obiegu']

                # Jeżeli jest to nowa doba, sprawdź który dzień w obiegu zaczyna się w stacji, w której kończyła jednostka dzień wczesniej
                if dzien_odjazdu != dzien_postoju:
                    if pomin_dzien_obiegu == dzien_w_obiegu:
                        continue
                    elif stacja_odjazdu != stacja_postoju:
                        pomin_dzien_obiegu = dzien_w_obiegu
                        continue
                    else:
                        pomin_dzien_obiegu = 0

                #  Jeżeli to ten sam dzień, sprawdź czy to poprawny dzień w obiegu
                else:
                    if dzien_w_obiegu != post_w_obiegu:
                        continue

                # Ten sam dzień i ten sam dzień_w_obiegu
                # Nowa doba, ale dzień w obiegu zaczyna się stacją kończoncą
                if stacja_odjazdu == stacja_postoju:
                    stacja_postoju = pociag['Rel. do']
                    postoj_od_godz = pociag['Prz. RT']
                    postoj_od_godz = self.sprawdz_format_godz(
                        postoj_od_godz)
                    dzien_postoju = pociag['Data']
                    post_w_obiegu = pociag['dzien_w_obiegu']

                    df_obieg_c.loc[i, 'nr_pojazdu'] = pojazd
                    df_obieg_c.loc[i, 'wykorzystano'] = 1
                    continue
                else:
                    continue

            df_obieg_c = df_obieg_c.sort_values(
                by=['nr_pojazdu', 'Data', 'dzien_w_obiegu', 'Odj. RT'])

        return df_obieg_c

    def sprawdz_format_godz(self, postoj_od_godz):
        if isinstance(postoj_od_godz, str):

            x = re.sub("^\[1\] ", "", postoj_od_godz)

            skladowe = re.split(":", x)

            godziny = int(skladowe[0])
            minuty = int(skladowe[1])

            time = 1 + (godziny/24)+(minuty/1440)

            return time
        else:
            return postoj_od_godz

    def dodatek_z_przebiegu(self, xw_book, df_przebieg_dla_obiegu, nr_obiegu, opis_obiegu):

        df = df_przebieg_dla_obiegu.copy()

        df_dates = df.drop_duplicates(subset="Data")

        df_dates = df_dates.set_index('Data')

        df_temp = pd.DataFrame()

        lista_df = []

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

                if df_temp_p.empty:
                    continue

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

        f_row = 2

        for i, df_u in enumerate(lista_df):

            df_u.loc[:, "opis_obiegu"] = opis_obiegu

            df_u.loc[:, "Uwagi"] = ''

            df_u.loc[:, "wariant_obiegu"] = i + 1

            ws_xl_dodatek_z_przebiegu = xw_book.sheets[f"obieg_{nr_obiegu}"]

            self.df_do_wykresow = pd.concat(
                [self.df_do_wykresow, df_u[self.df_do_wykresow.columns.intersection(df_u.columns)]], axis=0, ignore_index=True)

            df_u = df_u.loc[:, ["Odległość", "Nr poc.", "Rodz. poc.", "Rel. od",
                                "Odj. RT", "Rel. do", "Prz. RT", "Pojazdy", "Uwagi"]]

            # WKLEJANIE DO EXCELA

            if i == 0:
                ws_xl_dodatek_z_przebiegu[f"A{f_row}"].options(
                    pd.DataFrame, expand='table', index=False).value = df_u

                self.styl_tab(ws_xl_dodatek_z_przebiegu, f_row)

            else:

                f_row = f_row + l_row
                ws_xl_dodatek_z_przebiegu[f"A{f_row}"].expand('down').options(
                    pd.DataFrame, expand='table', index=False).value = df_u

                self.styl_tab(ws_xl_dodatek_z_przebiegu, f_row)

            kalendarz = fg()
            kalendarz.rysuj_kalendarz(
                daty_wariantu[i], df_dates, ws_xl_dodatek_z_przebiegu, f_row, 11)

            self.dodaj_opis(ws_xl_dodatek_z_przebiegu, f_row,
                            opis_obiegu, i, ilosc_pojazdow)

            df_u_len = len(df_u.index)

            if df_u_len < 14:
                l_row = 14
            else:
                l_row = df_u_len + 1

            # styl arkusza:
            self.styl_ark(ws_xl_dodatek_z_przebiegu)

    def rysuj_przebieg_do_xl(self, xw_book, df_przebieg_dla_obiegu, nr_obiegu):

        ws_xl_przebieg = xw_book.sheets[f"obieg_{nr_obiegu}"]

        ws_xl_przebieg["A1"].options(
            pd.DataFrame, expand='table', index=False).value = df_przebieg_dla_obiegu

        ws_xl_przebieg["J1"].expand("down").number_format = 'gg:mm'
        ws_xl_przebieg["L1"].expand("down").number_format = 'gg:mm'
        ws_xl_przebieg.autofit(axis="columns")

        # sprawdzenie czy w przebiegu pociąg kończy i zaczyna w tych samych stacjach
        for row in range(ws_xl_przebieg["A2"].expand("down").last_cell.row, 3, -1):
            st_o = ws_xl_przebieg[f"I{row}"].value
            st_p = ws_xl_przebieg[f"K{row - 1}"].value

            g_o = ws_xl_przebieg[f"J{row}"].value
            g_p = ws_xl_przebieg[f"L{row - 1}"].value

            data_o = ws_xl_przebieg[f"A{row}"].value
            data_p = ws_xl_przebieg[f"A{row - 1}"].value

            pojazd_o = ws_xl_przebieg[f"O{row}"].value
            pojazd_p = ws_xl_przebieg[f"O{row - 1}"].value

            if pojazd_o != pojazd_p:
                continue

            if st_o != st_p:
                ws_xl_przebieg.range(f"I{row}").color = (255, 0, 0)
                ws_xl_przebieg.range(f"K{row - 1}").color = (255, 0, 0)

            if data_o == data_p:
                if (g_o - g_p) < (10/1440):
                    ws_xl_przebieg.range(f"J{row}").color = (255, 0, 0)
                    ws_xl_przebieg.range(f"L{row - 1}").color = (255, 0, 0)
                elif (g_o - g_p) < (12/1440):
                    ws_xl_przebieg.range(f"J{row}").color = (255, 255, 9)
                    ws_xl_przebieg.range(f"L{row - 1}").color = (255, 255, 9)
                elif (g_o - g_p) < (14/1400):
                    ws_xl_przebieg.range(f"J{row}").color = (255, 250, 205)
                    ws_xl_przebieg.range(f"L{row - 1}").color = (255, 250, 205)

    def styl_tab(self, ws, start_row):
        zakres_dodatku = ws.range(f"A{start_row}").expand("table")

        # krawędzie wewnętrzne
        zakres_dodatku.api.Borders(11).Weight = 1

        zakres_dodatku.api.Borders(12).Weight = 1

        # Lewa krawędź
        zakres_dodatku.api.Borders(7).Weight = 3
        # Górna krawędź
        zakres_dodatku.api.Borders(8).Weight = 3
        # Dolna krawędź
        zakres_dodatku.api.Borders(9).Weight = 3
        # Prawa krawędź
        zakres_dodatku.api.Borders(10).Weight = 3

        # sprawdź czas przejścia
        for row in range(ws[f"A{start_row}"].expand("down").last_cell.row, start_row + 1, -1):
            godz_o = ws[f"E{row}"].value
            godz_p = ws[f"G{row - 1}"].value

            st_o = ws[f"D{row}"].value
            st_p = ws[f"F{row - 1}"].value

            if ((godz_o - godz_p) < 0) | (st_o != st_p):
                ws.range(f"A{row - 1}:I{row - 1}").api.Borders(9).Weight = 2
            elif ((godz_o - godz_p) < (10/1440)) & (st_o == st_p):
                ws.range(f"E{row}").color = (255, 0, 0)
                ws.range(f"G{row - 1}").color = (255, 0, 0)
            elif ((godz_o - godz_p) < (12/1440)) & (st_o == st_p):
                ws.range(f"E{row}").color = (255, 255, 9)
                ws.range(f"G{row - 1}").color = (255, 255, 9)
            elif ((godz_o - godz_p) < (14/1440)) & (st_o == st_p):
                ws.range(f"E{row}").color = (255, 250, 205)
                ws.range(f"G{row - 1}").color = (255, 250, 205)

    def styl_ark(self, ws):
        ws["E:E"].number_format = 'gg:mm'
        ws["G:G"].number_format = 'gg:mm'
        ws.autofit(axis="columns")
        ws["A:J"].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

    def dodaj_opis(self, ws, f_row, opis_obiegu, i, ilosc_pojazdow):

        dic_ilosci_pojazdow = {
            1: "jedno",
            2: "dwu",
            3: "trzy",
            4: "cztero",
            5: "pięcio",
            6: "sześcio",
            7: "siedmio",
            8: "ośmio",
            9: "dziewięcio"
        }

        ws.range(
            f"A{f_row - 1}").value = f"{i+1}. Obieg {dic_ilosci_pojazdow[ilosc_pojazdow]}dniowy: {opis_obiegu}_{i+1}"

    def wklej_df_do_wykr(self):

        wb_xl_wykresy = xw.Book(Path(__file__) / ".." / ".." /
                                "src" / "macros" / "wykresy_figurowe_baza.xlsm")

        ws_xl_tabela = wb_xl_wykresy.sheets['tabela']

        nr_wykresu = 0

        obieg_a = 0
        war_o_a = 0

        obieg_p = 0
        war_o_p = 0

        for i, row in self.df_do_wykresow.iterrows():
            obieg_a = row['opis_obiegu']
            war_o_a = row['wariant_obiegu']

            if obieg_a == obieg_p:
                if war_o_a == war_o_p:
                    self.df_do_wykresow.iloc[i, 0] = nr_wykresu
                else:
                    nr_wykresu += 1
                    self.df_do_wykresow.iloc[i, 0] = nr_wykresu
                    obieg_p = obieg_a
                    war_o_p = war_o_a
            else:
                nr_wykresu += 1
                self.df_do_wykresow.iloc[i, 0] = nr_wykresu
                obieg_p = obieg_a
                war_o_p = war_o_a

        self.df_do_wykresow = self.df_do_wykresow.iloc[:, :-1]

        ws_xl_tabela["B12"].options(
            pd.DataFrame, expand='table', index=False, header=False).value = self.df_do_wykresow

        wb_xl_wykresy.save(Path(__file__) / ".." / ".." /
                           "src" / "outputs" / "macros" / "wykresy_figurowe.xlsm")
