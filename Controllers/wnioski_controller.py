import xlwings as xw

from Models_xl.wnioski_model import Wnioski


class WnioskiController:

    def all(self):
        wnioski = Wnioski()
        xw.view(wnioski.all())

    def filtruj(self, wg_kolumny, wartosc):
        wnioski = Wnioski()
        xw.view(wnioski.filtruj(wg_kolumny, wartosc))
