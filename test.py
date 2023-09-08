import datetime
import os
from contextlib import suppress
from time import sleep

import pandas as pd

from config import mapping_path


df = pd.read_excel(mapping_path)

branches_to_execute = list(df[df['Сотрудник'] == 'Nusipova@magnum.kz']['Название филиала в Спруте'])

branches_to_execute1 = list(df[df['Сотрудник'] == 'Baishukova@magnum.kz']['Название филиала в Спруте'])

print(len(branches_to_execute), len(branches_to_execute1))



















