"""WMS 1.2.1 — a corpus of odd names for the text rules of WMS (the reading of a name by «Инструмент ЧПУ», «Метизы», «Поиск», the check of the names).

Why: the rules of 1.1.0–1.2.0 stopped on names nobody had thought of (a name that begins with a number: «10 мм сверло…», «М8 гайка»): in LibreOffice Basic «Or» and «And» evaluate both sides and
«IIf» both branches, so «Mid(s, 0, 1)» after a check «p = 1» still ran and raised «Invalid procedure call» — and one such name stopped the whole list. A test of the rules on a handful of
names written by hand cannot find that; this corpus can: every sample name with every beginning and every ending of it, with a character dropped or replaced, with a number or a sign put before
and after it, two names glued, and words of the domain taken at random (types, sizes, marks, standards, codes, threads, materials) — the same corpus for the same seed.

corpus(seed, n_random) → the list of names (no duplicates); sample(names, step) — every step-th one (a regular test takes a part, the whole is for a night run)."""
import random

from tool_names import TOOL_PARSE

# names of other goods of a warehouse: fasteners, electrics, chemicals, packing, short and degenerate ones
GOODS = [
    "Лампа H4 12V 60/55W", "Лампа H7 12V 55W", "Светодиод 3535 белый", "Кабель ВВГ 3х1,5", "Болт М8х30 DIN 933 8.8 оцинкованный", "Гайка М8 DIN 934", "Шайба 8 ГОСТ 11371",
    "Корпус фары левый", "Рассеиватель 150х220", "Отражатель Ø90", "Герметик силиконовый 310 мл", "Клей-герметик 2-х компонентный", "Подшипник 6204 2RS", "Резистор 10 кОм 0,25 Вт",
    "Конденсатор 100 мкФ 25 В", "Разъём 2-pin 2,54", "Провод ПВ-1 0,75 красный", "Скотч 48х50", "Лента двусторонняя 19 мм х 33 м", "Перчатки х/б размер 9", "Масло ИТД-68 20 л",
    "СОЖ эмульсол 20 л", "Труба стальная 25х2,5", "Лист 2х1250х2500", "Пруток Ø20 L=3000 сталь 45", "Заглушка 8 мм", "Винт M3x8 DIN 7985", "Саморез 4,2х32", "Шплинт 2х20",
    "Шуруп 4х40 по дереву", "Болт ГОСТ 7805-70 М10х50 класс 8.8", "Гайка самоконтрящаяся М8 ГОСТ 50271", "Шпилька М12х1,75х100 ГОСТ 22042", "Анкер клиновой 10х100", "Заклёпка 4,8х12 алюминий",
    "Фреза 3 зуба 12 мм", "Сверло 5 мм", "Метчик М6", "Цанга ER32 6-7", "3 фреза", "5 сверло", "12", "4z", "z4", "z", "зуб", "2зуб", "М", "м8", "d", "ø", "Ø", "-", "х", "×", "",
    "М8х30 болт", "М8 гайка", "8 мм болт", "10х1 труба", "6 мм шуруп", "5х40 саморез", "1 болт", "М 8", "M8", "8", "8х30", "10 мм сверло по металлу", "4z фреза концевая 10",
]

SPECIAL = ["¦", "\x01", "\x02", "\t", " ", "​", "\"", "'", "*", "?", "[", "]", "#", "=", "+", "-", "–", "—", "/", "\\", "(", ")", ".", ",", ";", ":", "%", "°", "№", "±", "×", "Ø", "ø", "⌀",
           "х", "x", "Х", "z", "Z", "d", "D", "м", "М", "0", "1", "9", "00", "99999999999999999999", "1,5", ".5", "5.", "1.2.3", "1,2,3", "--", "..", "  ", "&", "<", ">", "~", "^", "_", "|", "{", "}",
           "`", "@", "!", "$", "é", "ß", "ё", "Ё", "ю", "я", "’", "«", "»"]

PREFIXES = ["4 ", "10х", "2-х ", "0,5 ", ".5 ", "10z ", "6 мм ", "1 ", "00 ", "9999 ", "4z", "4зуб ", "3-х зубая ", "12 ", "8х1,25 ", "1/4 ", "ER32 ", "М8 ", "d12 ", "ø6 ", "Z4 ", "l=100 ", "№5 ", "(",
            "-", "х", " "]
SUFFIXES = [" 4", " 10х", " 4z", " z", " d", " ø", " М", " м", " мм", " 4 зуба", " 3-х", " Z=", " L=", " x", " х", " Ø", " 0", " 99999999999999999999", " DIN", " ГОСТ", " ISO", " ER", " BT", " SDS",
            " G1/4", " NPT", " ХВОСТОВИК", " хв."]

TYPES = ["фреза", "Фреза", "ФРЕЗА", "сверло", "Сверло спиральное", "метчик", "плашка", "развертка", "зенковка", "зенкер", "пластина", "державка", "патрон", "цанга", "оправка", "резец", "круг шлифовальный",
         "круг отрезной", "диск отрезной", "напильник", "полотно ленточное", "коронка", "бур", "борфреза", "гравер", "end mill", "drill", "tap", "insert", "расточная головка", "центр вращающийся",
         "пильный диск", "лепестковый круг", "СОЖ", "сопло плазменное", "болт", "гайка", "шайба", "винт", "саморез", "шуруп", "шпилька", "анкер", "заклепка", "труба", "лампа"]
KINDS = ["концевая", "спиральная", "шаровая", "твердосплавная", "быстрорежущая", "по алюминию", "по нержавейке", "торцевая", "дисковая", "резьбовая", "проходной", "отрезной", "машинно-ручной",
         "ступенчатое", "перовое", "центровочное", "удлиненное", "конусный", "оцинкованный", "нержавеющий", "самоконтрящаяся", "потайной", "с шестигранной головкой"]
NUMS = ["1", "2", "3", "4", "5", "6", "8", "10", "12", "16", "20", "25", "32", "50", "75", "100", "125", "150", "200", "0,5", "1,5", "2,5", "3,2", "8,5", "10.5", "0.8", "22,23", "1/4", "3/8", "1/2",
        "1 1/2", "4-20", "6-7", "32-40", "10..20", "10—20", "99999", "99999999999999999999", "0", "00", "007"]
SEPS = ["х", "x", "×", "*", "X", "Х", " х ", " x ", "-", "/", "–", ""]
MARKS = ["Ø", "ø", "⌀", "∅", "d", "D", "d=", "D=", "диам.", "диаметр", "dia", "L=", "l", "длина", "R", "r", "радиус", "Z", "z", "Z=", "z=", "хв.", "хвостовик", "shank", "№", "N", "n", "no", "угол",
         "шаг", "модуль", "М", "M", "мм", "mm"]
STDS = ["DIN 338", "DIN338", "ГОСТ 10902-77", "ГОСТ 10902", "ISO 2768", "ИСО 235", "ТУ 2-035-1", "ANSI B94", "JIS B4301", "DIN 371", "DIN 352", "DIN 1835", "ГОСТ 9140", "ГОСТ", "DIN", "ISO", "ту",
        "ГОСТ №", "DIN-338", "DIN 338-A", "ГОСТ 7798-70", "DIN 933", "ISO 4017", "DIN 934", "ГОСТ 5915"]
CODES = ["CNMG 120408", "TNMG160404", "DCMT 11T304", "APKT 1604", "MCLNR 2525M12", "SDJCR 1616H11", "WNMG 080408 PM 4325", "CCMT 09T304", "APMT1135", "XXXX 12", "ABCD 1234", "AB 12",
         "ABCDEFG 123456789"]
HOLD = ["ER32", "ER-32", "ER 32", "ER11", "BT40", "BT-40", "BT 30", "SK40", "SK 30", "HSK63", "HSK-A63", "HSK A 63", "HSK 63 A", "CAT40", "МК2", "МК 3", "КМ2", "NT40", "SDS-plus", "SDS plus", "SDS+",
        "SDS-max", "SDS", "конус Морзе 2", "Morse taper 3", "морзе", "ER", "BT", "SK", "HSK", "МК", "ER-", "BT-A"]
THR = ["М8х1,25", "M8x1.25", "М8", "М 8", "M8", "М10х1", "М6х1х30", "М8х30", "М8х1,25х30", "М100", "М101", "М0", "М", "М8х", "М8х1,25-6g", "G1/4", "G 1/4", "G1/4-19", "NPT 1/8", "NPT1/8",
       "UNC 1/4-20", "UNF 1/4-28", "Rc 3/8", "Rp 1/2", "BSP 1/2", "G", "R", "G1", "R1/2", "UNC", "NPT"]
MAT = ["Р6М5", "Р6М5К5", "ВК8", "ВК6", "HSS", "HSS-Co", "HSS-E", "HSSCo", "твердосплавная", "твердый сплав", "карбид", "vhm", "solid carbide", "Т15К6", "9ХС", "CBN", "алмаз", "быстрорез", "т/с",
       "керамика", "сталь 45", "А2", "А4", "8.8", "10.9", "12.9", "оцинкованный", "цинк", "латунь"]
COAT = ["TiN", "TiAlN", "AlTiN", "TiCN", "ZrN", "DLC", "Al2O3", "нитрид титана", "оксид", "паровая обработка", "без покрытия", "bright", "TiAlCrN"]
MISC = ["по металлу", "для станка", "с ЧПУ", "комплект 5 шт", "набор", "N5", "№ 7", "2 шт", "10 шт", "45 HRC", "HRC55", "8 HRC", "30°", "45 градусов", "90%", "1 кг", "(замена)", "[A]", "\"X\"",
        "'Y'", "новый", "б/у", "арт. 123", "арт.123456", "left", "правый", "левый", "RH", "LH", "L", "R", "К/Х", "Ц/Х", "ц/х", "к/х", "3-х зубая", "2-х зубая", "4 зуба", "2 флейты", "3 флейты",
        "6 канавок", "двухзаходная", "трёхзубая", "одно", "двух", "четыр", "пяти", "шести", "восьми", "М/Р", "м/р", "мр", "хвостовик 6", "хвостовик ø8", "хвост d=10", "цилиндрический хвостовик",
        "конический хвостовик"]
POOL = [TYPES, KINDS, NUMS, SEPS, MARKS, STDS, CODES, HOLD, THR, MAT, COAT, MISC, SPECIAL]
WEIGHT = [8, 4, 10, 3, 5, 3, 2, 3, 3, 3, 2, 4, 2]


def corpus(seed=1, n_random=12000):
    """the names (no duplicates) for the seed; n_random — how many of them are made of words taken at random"""
    rng = random.Random(seed)
    seeds = list(dict.fromkeys([n for n, _ in TOOL_PARSE] + GOODS))
    names = list(seeds)
    for s in seeds:
        for k in range(len(s) + 1):
            names.append(s[:k])
            names.append(s[k:])
        for k in range(len(s)):
            names.append(s[:k] + s[k + 1:])
        names += [p + s for p in PREFIXES]
        names += [s + x for x in SUFFIXES]
    for _ in range(2500):
        s = rng.choice(seeds)
        if s:
            k = rng.randrange(len(s))
            names.append(s[:k] + rng.choice(SPECIAL) + s[k:])
            k = rng.randrange(len(s))
            names.append(s[:k] + rng.choice(SPECIAL) + s[k + 1:])
    for _ in range(2000):
        names.append(rng.choice(seeds) + " " + rng.choice(seeds))
    for s in seeds[:400]:
        names.append(s.upper())
        names.append(s.lower())
    for _ in range(n_random):
        parts = [rng.choice(rng.choices(POOL, WEIGHT)[0]) for _ in range(rng.randint(1, 12))]
        nm = rng.choice([" ", " ", " ", "", "-", ", "]).join(parts)
        if rng.random() < 0.15:
            nm = nm.upper()
        elif rng.random() < 0.15:
            nm = nm.lower()
        if rng.random() < 0.3:
            nm = rng.choice(NUMS) + rng.choice([" ", "", "х", " мм ", "z ", "-х ", " зуб "]) + nm
        names.append(nm)
    # long and degenerate ones
    names += ["Фреза " + "1" * 300, "Фреза " + "9х" * 200, "Сверло " + "М8х1," * 100, "Фреза " + "z4 " * 100, "Сверло " + "1.2.3." * 100, "A" * 3000, "Фреза " + "ER32 " * 50, "Метчик " + "G1/4 " * 100,
              "Фреза " + "4 зуба " * 50, "х" * 500, "0" * 300, "Фреза " + "0" * 300, "Сверло " + "-" * 400, "Сверло " + "/" * 400, "4" * 100 + "z", "Сверло ø" + "9" * 400, "Болт " + "М8х" * 200,
              "8 мм " * 100 + "болт", "М8 " * 100 + "гайка"]
    return list(dict.fromkeys(n.replace("\n", " ") for n in names))


def sample(names, step=8, offset=0):
    """every step-th name (and the degenerate ones at the end — always)"""
    return names[offset::step] + names[-19:]
