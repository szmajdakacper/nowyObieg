"""
reports_concat
Skrypt, który skleja kilka wygenerowanych raportów w ISZTP
do jednego pliku : wnioski

"""
from pathlib import Path

import pandas as pd

from datetime import datetime

import pytz

raporty_dir = (Path(__file__) / ".." / ".." /
               "src" / "inputs" / "raporty").resolve()

wnioski_dir = (Path(__file__) / ".." / ".." /
               "src" / "outputs" / "wnioski").resolve()


# Getting the current date and time
dt = datetime.now(pytz.timezone('Europe/Vienna'))

# getting the timestamp
ts = dt.strftime("%d-%m-%Y__%H_%M")

parts = []

for path in (raporty_dir).rglob("*.xls*"):
    print(f"Reading: {path.name} ...")
    part = pd.read_excel(path, header=5, skipfooter=2, index_col="Lp")
    parts.append(part)

df = pd.concat(parts)

excel_output = f"{wnioski_dir}/wniosek{ts}.xlsx"

print(f"Saving: {excel_output} ...")
df.to_excel(excel_output)
