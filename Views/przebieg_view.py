from pathlib import Path

import xlwings as xw

from xlwings.utils import rgb_to_int

import pandas as pd

import datetime as dt

from datetime import datetime

import re

from env.env import FunkcjeGlobalne as fg

from Models_xl.obiegi_model import Obiegi


class PrzebiegView():

    def __init__(self):
        self.df_do_wykresow = pd.DataFrame(data=[], columns=["nr_wykresu", "Odległość", "Rodz. poc.", "Nr poc.", "Termin", "Uwagi", "Rel. od",
                                                             "Odj. RT", "Rel. do", "Prz. RT", "Pojazdy", "opis_obiegu", 'wariant_obiegu'])

    def przebieg_do_xl(self, df_przebieg):

        print("Tworzenie pliku Excel...")

        obiegi = Obiegi()
        obiegi = obiegi.all()

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

            opis_obiegu = obiegi.loc[nr_obiegu, "opis_obiegu"]

            self.dodatek_z_przebiegu(
                wb_xl_dodatek, df_przebieg_dla_obiegu, nr_obiegu, opis_obiegu)

        self.wklej_df_do_wykr()

        wb_xl_przebieg.save(Path(__file__) / ".." / ".." /
                            "src" / "outputs" / "pot" / "przebieg.xlsx")

        wb_xl_dodatek.save(Path(__file__) / ".." / ".." /
                           "src" / "outputs" / "pot" / "pot_rozszerzony.xlsx")

        print("Zakończono tworzenie pliku excel z sukcesem.")

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

            pojazd_o = ws_xl_przebieg[f"O{row}"].value
            pojazd_p = ws_xl_przebieg[f"O{row - 1}"].value

            if pojazd_o != pojazd_p:
                continue

            if st_o != st_p:
                ws_xl_przebieg.range(f"I{row}").color = (255, 0, 0)
                ws_xl_przebieg.range(f"K{row - 1}").color = (255, 0, 0)

    def rozpisz_przebieg_obiegu(self, df_przebieg, nr_obiegu, p_zakres, p_end):

        print(f"rysuje przebieg obiegu : {nr_obiegu}")

        df_przebieg_dla_obiegu = df_przebieg.loc[df_przebieg['nr_obiegu'] == nr_obiegu]

        df_przebieg_dla_obiegu = df_przebieg_dla_obiegu.sort_values(
            by=['Data', 'Odj. RT'])

        # jeżeli dzien w obiegu nie jest zdefiniowany to oznacz jako 1
        dfd = df_przebieg_dla_obiegu.copy()
        mask = dfd['dzien_w_obiegu'] == 0
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

                    # print('dnia:')
                    # print(p_dzien)

                    dfd = df_przebieg_dla_obiegu.copy()
                    mask = (dfd['Data'] == datetime.strftime(
                        p_dzien, '%Y-%m-%d')) & (dfd['dzien_w_obiegu'] == nr_dnia_obiegu)

                    if not dfd.loc[mask, 'wykorzystanie'].any() == 1:

                        dfd.loc[mask, 'nr_pojazdu'] = pojazd
                        dfd.loc[mask, 'wykorzystanie'] = 1
                        df_przebieg_dla_obiegu = dfd

                        # print('dla dnia w obiegu:')
                        # print(nr_dnia_obiegu)
                        # print('przypisuje pojazd:')
                        # print(pojazd)

                    # 1. nast_dzien_o następny dzień obiegu
                    nast_dzien_o = nr_dnia_obiegu + 1

                    if nast_dzien_o > ilosc_pojazdow:
                        nast_dzien_o = 1

                    if not p_dzien == p_end:

                        # sprawdź czy następnego dnia zaczyna w stacji, w której skończył
                        while not self.przejscie_nocne_pojazdu(
                                df_przebieg_dla_obiegu, p_dzien, nast_dzien_o, nr_dnia_obiegu, pojazd):

                            if nast_dzien_o == nr_dnia_obiegu:
                                if self.przejscie_nocne_pojazdu(
                                        df_przebieg_dla_obiegu, p_dzien, nast_dzien_o, nr_dnia_obiegu, pojazd) == False:
                                    print(
                                        f"BŁĄD_PRZEJŚCIA: Nie poprawne przejście jednoski z dnia {p_dzien} na dzień następny!")
                                    break

                            nast_dzien_o = nast_dzien_o + 1
                            if nast_dzien_o > ilosc_pojazdow:
                                nast_dzien_o = 1

                    nr_dnia_obiegu = nast_dzien_o
                    # print("następny numer obiegu to :")
                    # print(nr_dnia_obiegu)

        return df_przebieg_dla_obiegu

    def przejscie_nocne_pojazdu(self, df_przebieg_dla_obiegu, p_dzien, nr_dnia_obiegu, poprz_nr_dnia_ob, pojazd):

        # print(f"z nr: {poprz_nr_dnia_ob} przejscie na {nr_dnia_obiegu}")

        dfd = df_przebieg_dla_obiegu.copy()
        mask = (dfd['Data'] == datetime.strftime(
            p_dzien, '%Y-%m-%d')) & (dfd['dzien_w_obiegu'] == poprz_nr_dnia_ob)
        dfe = dfd.loc[mask, :]

        if dfe.empty:
            # print(f"brak pociągów dla {nr_dnia_obiegu} w {p_dzien}")
            # tego dnia żaden pociąg nie kursuje w obiegu, znajdź poprzedni dzien w którym kursował i sprawdź czy jest przejście
            wczesniejszy_dzien = p_dzien - dt.timedelta(days=1)
            while dfe.empty:
                mask = (dfd['Data'] == datetime.strftime(
                    wczesniejszy_dzien, '%Y-%m-%d')) & (dfd['nr_pojazdu'] == pojazd)
                dfe = dfd.loc[mask, :]
                wczesniejszy_dzien = wczesniejszy_dzien - dt.timedelta(days=1)

        ostatnia_stacja = dfe.iloc[-1, 9]

        nast_dzien = p_dzien + dt.timedelta(days=1)

        dff = df_przebieg_dla_obiegu.copy()
        mask = (dff['Data'] == datetime.strftime(
            (nast_dzien), '%Y-%m-%d')) & (dff['dzien_w_obiegu'] == nr_dnia_obiegu)
        dfg = dff.loc[mask, :]

        # jednostka zostaje w stacji przez całą dobę:
        if dfg.empty:
            print(
                f"Obieg pusty: dnia {datetime.strftime((p_dzien + dt.timedelta(days=1)), '%Y-%m-%d')}")

            return False

        try:
            pierwsza_nast_doba_stacja = dfg.iloc[0, 7]
            wykorzystanie = dfg.iloc[0, 13]
        except:
            print(
                f"Err: ost. s={ostatnia_stacja} na pierwsza_nast_doba_stacja Error dn. {datetime.strftime(p_dzien, '%Y-%m-%d')}")
            return False

        # print(f"przyj: {ostatnia_stacja} ; odj: {pierwsza_nast_doba_stacja}")

        if ostatnia_stacja != pierwsza_nast_doba_stacja:
            # print("'zwracam false'")
            return False
        else:
            if wykorzystanie == 1:
                # print("'zwracam false' bo wykorzystany")
                return False
            # print("zwracam true!")
            return True

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

            if ((godz_o - godz_p) < 0.00833) & (st_o == st_p):
                ws.range(f"E{row}").color = (255, 0, 0)
                ws.range(f"G{row - 1}").color = (255, 0, 0)
            elif ((godz_o - godz_p) < 0.01388) & (st_o == st_p):
                ws.range(f"E{row}").color = (255, 246, 204)
                ws.range(f"G{row - 1}").color = (255, 246, 204)

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

            df_obieg = self.rozpisz_pojazdy(df_obieg)

            # stwórz arkusz dla obiegu i zapisz przebieg do xl
            wb_xl_przebieg.sheets.add(f"obieg_{nr_obiegu}")
            ws_xl_przebieg_obiegu = wb_xl_przebieg.sheets[f"obieg_{nr_obiegu}"]

            ws_xl_przebieg_obiegu["A1"].expand('down').options(
                pd.DataFrame, expand='table', index=False).value = df_obieg

            # formatowanie arkusza
            self.rysuj_przebieg_do_xl(wb_xl_przebieg, df_obieg, nr_obiegu)

            # Stwórz dodatek (pot) rozszerzony z przebiegu
            self.dodatek_z_przebiegu(
                wb_xl_pot_roz, df_obieg, nr_obiegu, opis_obiegu)

        # Zapisz do plków Excela

        wb_xl_przebieg.save(Path(__file__) / ".." / ".." /
                            "src" / "outputs" / "pot" / "przebieg.xlsx")

        wb_xl_pot_roz.save(Path(__file__) / ".." / ".." /
                           "src" / "outputs" / "pot" / "pot_rozszerzony.xlsx")

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
            postoj_w_obiegu = ""
            pomin_dzien_obiegu = 0

            for i, pociag in df_obieg_c.iterrows():

                # Pomiń jeżeli pociąg jest w innym obiegu
                if pociag['wykorzystano'] == 1:
                    continue

                # Pierwszy obieg dla tego pojazdu
                if stacja_postoju == "":
                    stacja_postoju = pociag['Rel. do']
                    postoj_od_godz = pociag['Prz. RT']
                    postoj_od_godz = self.sprawdz_format_godz(postoj_od_godz)
                    dzien_postoju = pociag['Data']
                    postoj_w_obiegu = pociag['dzien_w_obiegu']

                    df_obieg_c.loc[i, 'nr_pojazdu'] = pojazd
                    df_obieg_c.loc[i, 'wykorzystano'] = 1
                    continue

                # Przypisz cechy pociągu do zmiennych
                stacja_odjazdu = pociag['Rel. od']
                godzina_odjazdu = pociag['Odj. RT']
                godzina_odjazdu = self.sprawdz_format_godz(godzina_odjazdu)
                dzien_odjazdu = pociag['Data']
                dzien_w_obiegu = pociag['dzien_w_obiegu']

                # Jeżeli jest to nowa doba, sprawdź który dzień w obiegu zaczyna się w stacji,w której kończyła jednostka dzień wczesniej
                if dzien_odjazdu != dzien_postoju:
                    if pomin_dzien_obiegu == dzien_w_obiegu:
                        continue
                    elif stacja_odjazdu != stacja_postoju:
                        pomin_dzien_obiegu = dzien_w_obiegu
                    else:
                        pomin_dzien_obiegu = 0

                if stacja_odjazdu == stacja_postoju:
                    if godzina_odjazdu > postoj_od_godz:

                        stacja_postoju = pociag['Rel. do']
                        postoj_od_godz = pociag['Prz. RT']
                        postoj_od_godz = self.sprawdz_format_godz(
                            postoj_od_godz)
                        dzien_postoju = pociag['Data']

                        df_obieg_c.loc[i, 'nr_pojazdu'] = pojazd
                        df_obieg_c.loc[i, 'wykorzystano'] = 1
                        continue

                    elif dzien_odjazdu > dzien_postoju:
                        stacja_postoju = pociag['Rel. do']
                        postoj_od_godz = pociag['Prz. RT']
                        postoj_od_godz = self.sprawdz_format_godz(
                            postoj_od_godz)
                        dzien_postoju = pociag['Data']

                        df_obieg_c.loc[i, 'nr_pojazdu'] = pojazd
                        df_obieg_c.loc[i, 'wykorzystano'] = 1
                        continue

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
