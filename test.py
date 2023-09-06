import datetime
import os
from contextlib import suppress
from time import sleep

import Levenshtein

branches = ['Алматинский филиал №1 ТОО "Magnum Cash&Carry"', 'Товарищество с ограниченной ответственностью Magnum Cash&Carry(777)', 'Алматинский филиал №2 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №3  ТОО "Magnum Cash&Carry"', 'Карагандинский Филиал №1 ТОО "Magnum Cash&Carry"', 'Филиал №1 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №4 ТОО "Magnum Cash&Carry" в г. Алматы', 'Филиал ТОО "Magnum Cash&Carry" №5 в г. Алматы', 'Алматинский филиал №6 ТОО "Magnum Cash&Carry"', 'Филиал Тест ТОО "Magnum cash&carry"', 'Алматинский филиал №7 ТОО "Magnum Cash&Carry"', 'Филиал ТОО "Magnum cash&carry" в г. Шымкент', 'Алматинский филиал №8 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №10 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №9 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №11 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №12 ТОО "Magnum Cash&Carry"', 'Филиал №2 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал ТОО "Magnum cash&carry" в г. Талдыкорган',
            'Алматинский филиал №14 ТОО "Magnum Cash&Carry"', 'Филиал №2 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №3 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №2 ТОО "Magnum Cash&Carry" в г.Талдыкорган', 'Алматинский филиал №16 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №15 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №17 ТОО "Magnum Cash&Carry"', 'Филиал №1 ТОО "Magnum Cash&Carry" в г.Каскелен', 'Алматинский филиал №20 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №18 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №19 ТОО "Magnum Cash&Carry"', 'Филиал №4 ТОО "Magnum Cash&Carry" в г.Шымкент', 'Карагандинский филиал №2 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №21 ТОО "Magnum Cash&Carry"', 'Филиал №1 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Алматинский филиал №22 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №23 ТОО "Magnum Cash&Carry"', 'Филиал №3 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Алматинский филиал №24 ТОО "Magnum Cash&Carry"',
            'Филиал №4 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №5 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Алматинский филиал №25 ТОО "Magnum Cash&Carry"', 'Филиал №1 в г. Кызылорда ТОО "Magnum Cash&Carry"', 'Алматинский филиал №26 ТОО "Magnum Cash&Carry"', 'Филиал №6 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №7 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №8 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №9 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №10 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №1 ТОО "Magnum Cash&Carry" в г. Тараз', 'Алматинский филиал №32 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №28 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №29 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №30 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №31 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №33 ТОО Magnum Cash&Carry', 'Алматинский филиал №34 ТОО Magnum Cash&Carry', 'Алматинский филиал №35 ТОО Magnum Cash&Carry',
            'Филиал №36 ТОО "Magnum Cash&Carry" в г Алматы',
            'Филиал №37 ТОО "Magnum Cash&Carry" в г. Алматы', 'Филиал №38 ТОО Magnum Cash&Carry в г. Алматы', 'Филиал №5 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №11 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №12 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №13 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №13 ТОО "Magnum Cash&Carry" в г.Алматы', 'Филиал №39 ТОО "Magnum Cash&Carry" в г.Алматы', 'Филиал №15 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №40 ТОО "MAGNUM CASH&CARRY" в г.Алматы', 'Алматинский филиал №41 ТОО "Magnum Cash&Carry"', 'Филиал №42 ТОО "Magnum Cash&Carry" в г.Алматы', 'Алматинский филиал №43 ТОО "Magnum Cash&Carry"', 'Филиал №14 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №6 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №7 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал РЦ №1 ТОО "Magnum Cash&Carry" в г.Астана', 'Филиал РЦ №2 ТОО "Magnum Cash&Carry" в г.Шымкент', 'Филиал №16 ТОО "MAGNUM CASH&CARRY" в г.Астана',
            'Филиал №17 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Карагандинский филиал №4 ТОО "Magnum Cash&Carry"', 'Карагандинский филиал №3 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №44 ТОО "Magnum Cash&Carry"', 'Филиал №8 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №2 ТОО "Magnum Cash&Carry" в г. Тараз', 'Карагандинский филиал №5 ТОО "Magnum Cash&Carry"', 'Филиал №45 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 ТОО "Magnum Cash&Carry" в г.Есик', 'Филиал №19 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №46 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №24 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Алматинский филиал №49 ТОО "Magnum Cash&Carry"', 'Филиал №21 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №9 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №48 ТОО «MAGNUM СASH&CARRY» в г.Алматы', 'Филиал №10 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №20 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №56 ТОО «MAGNUM CASH&CARRY» в г. Алматы',
            'Филиал №28 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №50 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №53 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 ТОО «МAGNUM СASH&CARRY» в г. Туркестан', 'Филиал №22 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №7 ТОО «МAGNUM СASH&CARRY» в г.Караганда', 'Филиал №51 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №23 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №18 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №52 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №25 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №54 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №55 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №26 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №27 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №29 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №30 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №60 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №2 в г. Кызылорда ТОО "Magnum Cash&Carry"',
            'Карагандинский филиал №6 ТОО "Magnum Cash&Carry"', 'Дискаунтер Реалист №11', 'Филиал №59 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №58 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 ТОО "MAGNUM CASH&CARRY"  в г. Усть-Каменогорск', 'Филиал №2 ТОО "MAGNUM CASH&CARRY"  в г. Усть-Каменогорск', 'Филиал №31 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'ДУЦП ТОО «Magnum Cash&Carry»', 'Филиал №33 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №35 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №32 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №41 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №34 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №36 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №37 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №2 ТОО "Magnum Cash&Carry" в г.Каскелен', 'Филиал №47 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №2 ТОО «МAGNUM СASH&CARRY» в г. Туркестан', 'Филиал №61 ТОО «MAGNUM CASH&CARRY» в г. Алматы',
            'Филиал №38 ТОО "MAGNUM CASH&CARRY" в г.Астана',
            'Филиал №39 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №40 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №42 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №51 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №48 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №49 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №43 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №44 ТОО "MAGNUM CASH&CARRY" г.Астана', 'Филиал №53 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №45 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №57 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №46 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №47 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №50 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №52 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №11 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №56 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №54 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №55 ТОО "MAGNUM CASH&CARRY" в г.Астана',
            'Филиал №62 ТОО «MAGNUM CASH&CARRY» в г. Алматы',
            'Филиал №63 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №12 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №68 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №3 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №14 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №67 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Распределительный центр №3 в Алматинской области', 'Филиал №66 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №69 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №63 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №64 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №57 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №62 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №15 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Алматинский филиал №71 ТОО "Magnum Cash&Carry"', 'Филиал №20 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №17 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №73 ТОО «MAGNUM СASH&CARRY» в г. Алматы', 'Филиал №72 ТОО «MAGNUM СASH&CARRY» в г. Алматы',
            'Филиал №18 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №19 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №65 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №3 ТОО «МAGNUM СASH&CARRY» по Туркестанской области', 'Филиал №61 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №20 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №21 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №58 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №1 ТОО "Magnum Cash&Carry" в г.Конаев', 'Филиал №3 ТОО "MAGNUM CASH&CARRY"  в г. Усть-Каменогорск', 'Филиал №19 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №22 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №64 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №65 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №17 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал РЦ №4 ТОО "Magnum Cash&Carry" в г.Петропавловск', 'Филиал №2 ТОО "Magnum Cash&Carry" в г. Петропавловск',
            'Филиал №11 ТОО "Magnum Cash&Carry" в г. Петропавловск',
            'Филиал №4 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №5 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №15 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №7 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №8 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №6 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №13 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №18 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №10 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №12 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №9 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №16 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №14 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №3 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №59 ТОО «MAGNUM CASH&CARRY» в г.Астана', 'Филиал №13 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №75 ТОО "Magnum Сash&Сarry" в г. Алматы',
            'Филиал №60 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №21 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №22 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №23 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №70 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №24 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №25 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №26 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №27 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №28 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №29 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №30 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №31 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №32 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №33 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №34 ТОО "Magnum Cash&Carry" в г. Шымкент', 'Филиал №4 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №5 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №6 ТОО "Magnum Cash&Carry" в г. Тараз',
            'Филиал №7 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №8 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №9 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №10 ТОО "Magnum Cash&Carry" в г. Тараз', 'Филиал №66 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №67 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №35 ТОО «MAGNUM СASH&CARRY» в г. Шымкент', 'Филиал №23 ТОО "Magnum Cash&Carry" в г. Петропавловск', 'Филиал №76 ТОО «MAGNUM CASH&CARRY» в г. Алматы', 'Филиал №1 Маркет холл ТОО "Magnum Cash&Carry" в г. Алматы', 'Филиал №68 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №69 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №71 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №73 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №70 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №74 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №72 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №75 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №77 ТОО "Magnum Сash&Сarry" в г. Алматы',
            'Филиал №78 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №76 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №77 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №79 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Алматинский филиал №80 ТОО "Magnum Cash&Carry"', 'Алматинский филиал №81 ТОО "Magnum Cash&Carry"', 'Филиал №82 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №83 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №4 ТОО «МAGNUM СASH&CARRY» в г. Туркестан', 'Филиал №79 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №84 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №85 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №80 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №81 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №86 ТОО "Magnum Сash&Сarry" в г. Алматы', 'Филиал №82 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №83 ТОО "MAGNUM CASH&CARRY" в г.Астана', 'Филиал №74 ТОО "Magnum Сash&Сarry" в г. Алматы'
]


def get_branches_to_execute(df1, branches_with_quote):

    def apply_replacements(cell):
        with suppress(Exception):
            cell_ = cell.replace('.', '').replace('"', '')
            # print(cell_)
            return cell.replace('.', '').replace('"', '')
        return cell

    branches_without_quote = pd.DataFrame(branches_with_quote)
    branches_without_quote = branches_without_quote.applymap(apply_replacements)
    branches_without_quote.columns = ['stores']

    def apply_replacements1(cell):
        with suppress(Exception):
            cell_ = cell.replace('.xlsx', '')
            # print(cell_)
            return cell.replace('.xlsx', '')
        return cell

    df1.columns = ['№', 'Короткое название филиала', 'store_name', 'Сотрудник']
    df1['store_name'] = df1['store_name'].astype(str)
    df1['store_name'] = df1['store_name'].apply(apply_replacements1)
    # print(df1['store_name'])

    import numpy as np

    # print(f"{len(np.setdiff1d(np.asarray(branches_without_quote['stores']), np.asarray(df1['store_name'])))}")

    skipped_branches = np.setdiff1d(np.asarray(branches_without_quote['stores']), np.asarray(df1['store_name']))
    # print(skipped_branches)
    # print('-------------------------------------------------------------------------')

    branches_to_execute_ = []

    for branch in branches_with_quote:
        branch_ = str(branch).lower().replace('.', '')

        for branch1 in skipped_branches:
            branch1_ = str(branch1).lower()

            diff = Levenshtein.distance(branch_, branch1_)
            if diff <= 2:
                # print(f'TO EXECUTE: {diff} | {branch}, {branch1}')
                branches_to_execute_.append(branch)
                break

    for br in branches_to_execute_:
        print(br)
    print(len(branches_to_execute_))
    return branches_to_execute_


import pandas as pd

from config import mapping_path

df = pd.read_excel(mapping_path)

print(len(df), len(df) - len(df[df['Сотрудник'] == 'Baishukova@magnum.kz']), len(df) - len(df[df['Сотрудник'] == 'Nusipova@magnum.kz']))

branches = ['АФ1', 'АФ10', 'АФ11', 'АФ12', 'АФ14', 'АФ15', 'АФ16', 'АФ17', 'АФ18', 'АФ19', 'АФ2', 'АФ20', 'АФ21', 'АФ22', 'АФ23', 'АФ24', 'АФ25', 'АФ26', 'АФ28', 'АФ29', 'АФ3', 'АФ30', 'АФ31', 'АФ32', 'АФ33', 'АФ34', 'АФ35', 'АФ41', 'АФ43', 'АФ44', 'АФ49', 'АФ6', 'АФ7', 'АФ71', 'АФ8', 'АФ80', 'АФ9', 'КФ1', 'КФ2', 'КФ3', 'КФ4', 'ДР3 (КФ5)', 'КФ6', 'ТКФ1', 'ШФ1', 'КЗФ1', 'УКФ1', 'ППФ1', 'ТЗФ1', 'АСФ1', 'ЕКФ1', 'ФКС1', 'КПФ1', 'ТФ1', 'ППФ10', 'ТЗФ10', 'АСФ10', 'ШФ10', 'ППФ11', 'ШФ11', 'АСФ11', 'ППФ12', 'ШФ12', 'АСФ12', 'ППФ13', 'ШФ13', 'АСФ13', 'ППФ14', 'ШФ14', 'АСФ14', 'ППФ15', 'ШФ15', 'АСФ15', 'ППФ16', 'АСФ16', 'ППФ17', 'АСФ17', 'ШФ17', 'ППФ18', 'АСФ18', 'ШФ18', 'ППФ19', 'АСФ19', 'ШФ19', 'КЗФ2', 'УКФ2', 'ППФ2', 'ТЗФ2', 'ШФ2', 'АСФ2', 'ФКС2', 'ТКФ2', 'ТФ2', 'ППФ20', 'ШФ20', 'АСФ20', 'ППФ21', 'ШФ21', 'АСФ21', 'ППФ22', 'ШФ22', 'АСФ22', 'ШФ23', 'АСФ23', 'ШФ24', 'АСФ24', 'ШФ25', 'АСФ25', 'ШФ26', 'АСФ26', 'ШФ27', 'АСФ27', 'ШФ28', 'АСФ28(др8)', 'ШФ29', 'АСФ29', 'УКФ3', 'ППФ3', 'ТЗФ3', 'ШФ3', 'АСФ3', 'ТФ3', 'ШФ30', 'АСФ30', 'ШФ31', 'АСФ31', 'ШФ32', 'АСФ32', 'ШФ33', 'АСФ33', 'ШФ34', 'АСФ34', 'АСФ35', 'ШФ35 ОПТ', 'АФ36', 'АСФ36', 'АФ37', 'АСФ37', 'АФ38', 'АСФ38', 'АФ39', 'АСФ39', 'АФ4', 'ППФ4', 'ТЗФ4', 'АСФ4', 'ШФ4', 'ТФ4', 'АФ40', 'АСФ40', 'АСФ41', 'АФ42', 'АСФ42', 'АСФ43', 'АСФ44', 'АСФ45', 'АФ45(др4)', 'АСФ46', 'АФ46', 'АСФ47', 'АФ47', 'АСФ48', 'АФ48', 'ППФ5', 'ТЗФ5', 'ШФ5', 'АСФ5', 'АФ50', 'АСФ50', 'АСФ51', 'АФ51', 'АСФ52', 'АФ52', 'АСФ53', 'АФ53', 'АСФ54', 'АФ54', 'АСФ55', 'АФ55', 'АФ56(др9)', 'АСФ56', 'АСФ57', 'АФ57', 'АСФ58', 'АФ58', 'АФ59', 'АСФ59', 'ППФ6', 'ТЗФ6', 'ШФ6', 'АСФ6', 'АСФ60', 'АФ60(др10)', 'АСФ61', 'АФ61', 'АСФ62', 'АФ62', 'АСФ63', 'АФ63', 'АСФ64', 'АФ64', 'АСФ65', 'АФ65', 'АСФ66', 'АФ66', 'АСФ67', 'АФ67', 'АСФ68', 'АФ68', 'АСФ69', 'АФ69', 'ППФ7', 'ТЗФ7', 'ШФ7', 'АСФ7', 'КФ7', 'АСФ70', 'АФ70', 'АСФ71', 'АСФ72', 'АФ72', 'АСФ73', 'АФ73', 'АСФ74', 'АСФ75', 'АФ75', 'АСФ76', 'АФ76', 'АСФ77', 'АФ77', 'АФ78', 'АСФ79', 'АФ79', 'ППФ8', 'ТЗФ8', 'ШФ8', 'АСФ8', 'АСФ80', 'АСФ81', 'АСФ82', 'АФ82', 'АСФ83', 'АФ83', 'АФ84', 'АФ86', 'ППФ9', 'ТЗФ9', 'АСФ9', 'ШФ9']

# for i in range(len(branches)):
#     branches[i] = branches[i].replace('.xlsx', '')
print(list(df[df['Сотрудник'] == 'Nusipova@magnum.kz']['Название филиала в Спруте']))
# c = 0
# for i in range(len(df[df['Сотрудник'] == 'Baishukova@magnum.kz'])):
#     print(df['Короткое название филиала'].iloc[i])
#     if df['Короткое название филиала'].iloc[i] in branches:
#         c += 1
#         print(df['Название филиала в Спруте'].iloc[i])
#         print('-------------------')
#     else:
#         pass
#         # print(i)
#
# print(c)
#
# print(len(df[df['Сотрудник'] == 'Baishukova@magnum.kz']['Название филиала в Спруте']))
c = 0
# for i in df[df['Сотрудник'] == 'Nusipova@magnum.kz']['Название филиала в Спруте']:
#     print(i)
#     if i in branches:
#         c += 1
#         # print(i)
#     else:
#         print(i)
#
# print(c)
#
# print(len(df[df['Сотрудник'] == 'Nusipova@magnum.kz']['Название филиала в Спруте']))













































