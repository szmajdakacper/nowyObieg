from Controllers.dodatek_controller import DodatekController
from Views.dodatek_view import DodatekView

df_dodatek = DodatekController()

xl_dodatek = DodatekView()

xl_dodatek.dodatek_do_xl(df_dodatek.stworz_dodatek())
