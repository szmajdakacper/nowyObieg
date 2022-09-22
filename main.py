from datetime import datetime

from Controllers.dodatek_controller import DodatekController
from Views.dodatek_view import DodatekView
from Views.przebieg_view import PrzebiegView
from Controllers.przebieg_controller import PrzebiegController

start = datetime.now()

df_dodatek = DodatekController()

xl_dodatek = DodatekView()

xl_dodatek.dodatek_do_xl(df_dodatek.stworz_dodatek())

start_przebiegu = "06-11-2022"
koniec_przebiegu = "10-12-2022"

przebieg = PrzebiegController(start_przebiegu, koniec_przebiegu)

xl_przebieg = PrzebiegView()

xl_przebieg.przebieg_do_xl(przebieg.stworz_przebieg())

end = datetime.now()

diff = end - start

print(f"Zrobienie dodatków zajęło: {int(diff.total_seconds())}s")
