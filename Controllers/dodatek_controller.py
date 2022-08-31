import pandas as pd

from Models_xl.wnioski_model import Wnioski

from Models_xl.obiegi_model import Obiegi

from Models_xl.pot_model import Pot


class DodatekController:

    def stworz_dodatek(self):

        # stwórz tablicę DataFrame dodatku
        df_dodatek = pd.DataFrame(data=[], columns=['rel_w_pot', 'Nr gr. poc.', 'nr_obiegu', 'opis_obiegu', 'Odległość', 'Rodz. poc.',
                                  'Nr poc.', 'Termin', 'Uwagi', 'Rel. od', 'Odj. RT', 'Rel. do', 'Prz. RT', 'Zestawienie'])

        # stwórz obiekt modelu Wnioski
        wnioski = Wnioski()

        # stwórz obiekt modelu plan obiegów taboru (Pot)
        pot = Pot()
        pot_all = pot.all()

        # pobierz wszystkie zdefiniowane obiegi
        obiegi = Obiegi()

        obiegi = obiegi.all()

        # Pętla 1. Po wszystkich obiegach, dla każdego pociągu -----------------------------------------

        for nr_obiegu, obieg in obiegi.iterrows():

            print(f"Analiza obiegu: {int(nr_obiegu)}. {obieg['opis_obiegu']}")
            pociagi_w_obiegu = pot.filtruj("nr_obiegu", nr_obiegu)

            # sprawdź każdy pociąg w obiegu, zdefiniowany w pliku inputs/plan_obiegów_taboru.xlsx
            for i, poc_w_obiegu in pociagi_w_obiegu.iterrows():
                rel_w_pot = i
                print(
                    f"{rel_w_pot}. Pociąg_id: {int(poc_w_obiegu['Nr gr. poc.'])}")

                try:
                    # pobierz wnioski dla pociągu po jego id (nr gr. poc.), tylko z unikatowymi godzinami
                    wnioski_dla_poc_id = wnioski.pobierz_do_dodatku(
                        "Nr gr. poc.", int(poc_w_obiegu['Nr gr. poc.']), nr_obiegu, rel_w_pot)
                except:
                    print(
                        f"Brak wnioski o nr id: {poc_w_obiegu['Nr gr. poc.']}, lub pobieranie wywołało błąd.")
                    continue

                # dodaj do dataframe df_dodatek kolejny pociąg (z ew. wariantami)
                df_dodatek = pd.concat(
                    [df_dodatek, wnioski_dla_poc_id], axis=0, ignore_index=True)

        # Pętla 2. Po wszytskich relacjach w dodatku, dopisz dodatkowe informację ---------------------

        for nr_rel_w_dodatku, rel_w_dodatku in df_dodatek.iterrows():

            try:
                # opis obiegu:
                opis_obiegu = obiegi.loc[rel_w_dodatku["nr_obiegu"],
                                         "opis_obiegu"]
                df_dodatek.loc[nr_rel_w_dodatku, "opis_obiegu"] = opis_obiegu

                # zestawienie:
                zestawienie = obiegi.loc[rel_w_dodatku["nr_obiegu"],
                                         "Zestawienie"]
                df_dodatek.loc[nr_rel_w_dodatku, "Zestawienie"] = zestawienie

                # termin:
                termin = pot_all.loc[rel_w_dodatku["rel_w_pot"], "Termin"]
                df_dodatek.loc[nr_rel_w_dodatku, "Termin"] = termin

                # uwagi:
                uwagi = pot_all.loc[rel_w_dodatku["rel_w_pot"], "Uwagi"]
                df_dodatek.loc[nr_rel_w_dodatku, "Uwagi"] = uwagi
            except Exception as e:
                print(
                    f"Próba dodania dodatkowych informacji dla {nr_rel_w_dodatku} zakończyła się błędem. Sprawdź to.")
                print(opis_obiegu)
                continue

        # Zwróć dodatek jako obiekt DateFrame ---------------------------------------------------------

        return df_dodatek
