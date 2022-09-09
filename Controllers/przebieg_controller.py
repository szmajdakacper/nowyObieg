from datetime import datetime
import pandas as pd
import re

from Models_xl.obiegi_model import Obiegi
from Models_xl.pot_model import Pot
from Models_xl.wnioski_model import Wnioski

from env.env import ZmienneSrodowiskowe


class PrzebiegController():

    def __init__(self, start_date, end_date):
        self.start_date = datetime.strptime(start_date, '%d-%m-%Y')
        self.end_date = datetime.strptime(end_date, '%d-%m-%Y')
        self.zmienneSrodowiskowe = ZmienneSrodowiskowe()

    def stworz_przebieg(self):
        start_date = self.start_date
        end_date = self.end_date

        df_przebieg = pd.DataFrame(data=[], columns=['Data', 'nr_obiegu', 'dzien_w_obiegu', 'Odległość',
                                   'Rodz. poc.', 'Nr poc.', 'Termin', 'Rel. od', 'Odj. RT', 'Rel. do', 'Prz. RT', 'Pojazdy'])

        # pobierz informacje z plików Excela
        # stwórz obiekt modelu Wnioski
        wnioski = Wnioski()
        wnioski_all = wnioski.all()

        # stwórz obiekt modelu plan obiegów taboru (Pot)
        pot = Pot()
        pot_all = pot.all()

        # pobierz wszystkie zdefiniowane obiegi
        obiegi = Obiegi()
        obiegi_all = obiegi.all()

        print(
            f"Rozpoczęto przebieg w zakresie dat: {start_date.strftime('%d-%m-%Y')} do {end_date.strftime('%d-%m-%Y')}")

        zakres_dat = pd.date_range(start_date, end_date)

        # Pętla po każdym obiegu
        for nr_obiegu, obieg in obiegi_all.iterrows():
            pociagi_w_obiegu = pot.filtruj("nr_obiegu", nr_obiegu)
            print(f"Analiza obiegu: {nr_obiegu}")

            # Sprawdź każdy dzień w zakresie dat
            for dzien in zakres_dat:
                # print(f"Dzień: {dzien.strftime('%d.%m')}:")

                # wyfiltruj wnioski, które tego dnia kursują
                kursujace_zamowienia = wnioski.filtruj_daty_kursowania(
                    dzien.strftime('%Y.%m.%d'))

                # Pętla po każdym pociągu w obiegu
                for i, poc_w_obiegu in pociagi_w_obiegu.iterrows():

                    # Sprawdź czy pociąg w tym dniu kursuje w tym obiegu (po zadanym terminie)
                    if poc_w_obiegu["Termin"] == "H":
                        # codziennie
                        pass

                    elif poc_w_obiegu["Termin"] == "D":
                        # w dni robocze
                        if dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela:
                            continue
                        elif (dzien.isoweekday() == 6) | (dzien.isoweekday() == 7):
                            continue

                    elif poc_w_obiegu["Termin"] == "C":
                        # w weekendy i święta
                        if dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela:
                            pass
                        elif (dzien.isoweekday() == 6) | (dzien.isoweekday() == 7):
                            pass
                        else:
                            continue

                    elif poc_w_obiegu["Termin"] == "A":
                        # od poniedziałku do piątku
                        if (dzien.isoweekday() == 6) | (dzien.isoweekday() == 7):
                            continue

                    elif poc_w_obiegu["Termin"] == "B":
                        # bez soboty
                        if (dzien.isoweekday() == 6):
                            continue

                    elif poc_w_obiegu["Termin"] == "E":
                        # od poniedziałku do soboty oprócz świąt
                        if dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela:
                            continue
                        elif (dzien.isoweekday() == 7):
                            continue

                    elif re.search(r"^\[\d-\d\]$", poc_w_obiegu["Termin"]):
                        # dni tygodnia od - do
                        od_do = re.findall(r"\d", poc_w_obiegu["Termin"])
                        if (dzien.isoweekday() >= int(od_do[0])) & (dzien.isoweekday() <= int(od_do[1])):
                            pass
                        else:
                            continue

                    elif re.search(r"^\[\d\]$", poc_w_obiegu["Termin"]):
                        # jeden dzien w tygodniu
                        dzien_w_tyg = re.findall(r"\d", poc_w_obiegu["Termin"])
                        if dzien.isoweekday() == int(dzien_w_tyg[0]):
                            pass
                        else:
                            continue

                    elif re.search(r"^\[\d\]\+$", poc_w_obiegu["Termin"]):
                        # jeden dzien w tygodniu
                        dzien_w_tyg = re.findall(r"\d", poc_w_obiegu["Termin"])
                        if (dzien.isoweekday() == int(dzien_w_tyg[0])) | (dzien in self.zmienneSrodowiskowe.swieta_poza_niedziela):
                            pass
                        else:
                            continue

                    else:
                        print(
                            f"brak rozważanego terminu kursowania {poc_w_obiegu['Termin']}")

                    df_lista_zamowien = wnioski.pobierz_liste_zamowien(
                        poc_w_obiegu["Nr gr. poc."])

                    # print(f"Pociąg id: {poc_w_obiegu['Nr gr. poc.']}")

                    # Znajdź zamówienie które danego dnia kursuje
                    for id_zamowienia, zamowienie in df_lista_zamowien.iterrows():
                        znalezione_zamowienie = kursujace_zamowienia.loc[kursujace_zamowienia['Nr zam.']
                                                                         == zamowienie['Nr zam.']]
                        if not znalezione_zamowienie.empty:
                            wniosek_poc = wnioski.filtruj(
                                "Nr zam.", zamowienie["Nr zam."])

                            wniosek_poc.insert(
                                0, "Data", dzien.strftime('%Y-%m-%d'))

                            wniosek_poc.insert(1, "nr_obiegu", nr_obiegu)

                            wniosek_poc.insert(
                                2, "dzien_w_obiegu", poc_w_obiegu["dzien_w_obiegu"])

                            wniosek_poc.insert(
                                3, "Termin", poc_w_obiegu["Termin"])

                            wniosek_poc = wniosek_poc.loc[:, ['Data', 'nr_obiegu', 'dzien_w_obiegu', 'Odległość', 'Rodz. poc.',
                                                              'Nr poc.', 'Termin', 'Rel. od', 'Odj. RT', 'Rel. do', 'Prz. RT', 'Pojazdy']]

                            df_przebieg = pd.concat(
                                [df_przebieg, wniosek_poc], axis=0, ignore_index=True)

        return df_przebieg
