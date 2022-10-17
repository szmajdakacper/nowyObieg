from datetime import datetime

from Controllers.dodatek_controller import DodatekController
from Views.dodatek_view import DodatekView
from Views.przebieg_view import PrzebiegView
from Controllers.przebieg_controller import PrzebiegController

start = datetime.now()


# # -- Dodatek:

df_dodatek = DodatekController()

xl_dodatek = DodatekView()

xl_dodatek.dodatek_do_xl(df_dodatek.stworz_dodatek())


# -- Przebieg obiegów:

start_przebiegu = "11-12-2022"
koniec_przebiegu = "11-03-2023"

przebieg = PrzebiegController(start_przebiegu, koniec_przebiegu)

xl_przebieg = PrzebiegView()

xl_przebieg.pokaz_przebieg_df(przebieg.stworz_przebieg())


end = datetime.now()

diff = end - start

print(f"Zrobienie dodatków zajęło: {int(diff.total_seconds())}s")
