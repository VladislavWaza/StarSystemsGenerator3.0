from openpyxl import Workbook, load_workbook
import random
import math
rand = 0
TypesStar = ['Горячая звезда главной последовательности', 'Горячий сверхгигант главной последовательности',
            'Звезда главной последовательности', 'Гигант главной последовательности',
            'Сверхгигант главной последовательности', 'Коричневый карлик', 'Белый карлик', 'Чёрный карлик', 'Пульсар',
            'Магнитар']


def CountStar():
    global rand
    if rand == -1:
        inp = -1
    else:
        inp = int(input("Сколько звезд в системе? "))
        if inp == -2:
            rand = -1
            inp = -1
    if inp == -1:
        inp = random.randint(1, 1000)
        if inp >= 999:
            return 8
        elif inp >= 995:
            return 7
        elif inp >= 990:
            return 6
        elif inp >= 980:
            return 5
        elif inp >= 960:
            return 4
        elif inp >= 900:
            return 3
        elif inp >= 800:
            return 2
        else:
            return 1
    else:
        return inp

def TypeStar(name_star):
    global TypesStar, rand
    if rand == -1:
        inp = -1
    else:
        print("Выберете тип " + name_star)
        for _ in range(10):
            print(str(_) + '.' + TypesStar[_ - 1])
        inp = int(input())
        if inp == -2:
            rand = -1
            inp = -1
    if inp == -1:
        inp = random.randint(1, 100)
        if inp >= 96:
            return 1
        elif inp >= 95:
            return 2
        elif inp >= 21:
            return 3
        elif inp >= 17:
            return 4
        elif inp >= 15:
            return 5
        elif inp >= 10:
            return 6
        elif inp >= 5:
            return 7
        elif inp >= 4:
            return 8
        elif inp >= 2:
            return 9
        else:
            return 10
    else:
        return inp

def CountPlanet(name_star):
    global rand
    if rand == -1:
        inp = -1
    else:
        inp = int(input("Сколько планет у звезды "+name_star+"? "))
        if inp == -2:
            rand = -1
            inp = -1
    if inp == -1:
        inp = random.randint(1, 31)
        if inp == 31:
            inp = 0
        elif inp == 30:
            inp = 1
        elif inp == 29:
            inp = 2
        elif inp == 28 or inp == 27:
            inp = 3
        elif inp == 26 or inp == 25:
            inp = 4
        elif inp == 24 or inp == 23 or inp == 22:
            inp = 5
        elif inp == 21 or inp == 20 or inp == 19:
            inp = 6
        elif inp == 18 or inp == 17 or inp == 16:
            inp = 7
        elif inp == 15 or inp == 14:
            inp = 8
        elif inp == 13 or inp == 12:
            inp = 9
        elif inp == 11:
            inp = 10
        elif inp == 10:
            inp = 11
        elif inp == 9:
            inp = 12
        elif inp == 8:
            inp = 13
        elif inp == 7:
            inp = 14
        elif inp == 6:
            inp = 15
        elif inp == 5:
            inp = 16
        elif inp == 4:
            inp = 17
        elif inp == 3:
            inp = 18
        elif inp == 2:
            inp = 19
        elif inp == 1:
            inp = 20
    return inp

def KlassStar(name_star, type_star):
    global rand
    if rand == -1:
        inp = '-1'
    else:
        inp = 0
    if type_star == 1:
        if inp != '-1':
            print("Выберете класс "+name_star)
            print('O\nB\nA\nF')
            inp = input()
        if inp == '-2':
            rand = -1
            inp = '-1'
        if inp == '-1':
            inp = random.randint(1, 100)
            if inp >= 30:
                return 'F'
            elif inp >= 7:
                return 'A'
            elif inp >= 3:
                return 'B'
            else:
                return 'O'
        else:
            return inp
    elif type_star == 2:
        if inp != '-1':
            print("Выберете класс " + name_star)
            print('Oc\nBc\nAc\nFc')
            inp = input()
        if inp == '-2':
            rand = -1
            inp = '-1'
        if inp == '-1':
            inp = random.randint(1, 100)
            if inp >= 30:
                return 'Fс'
            elif inp >= 7:
                return 'Aс'
            elif inp >= 3:
                return 'Bс'
            else:
                return 'Oс'
        else:
            return inp
    elif type_star == 3:
        if inp != '-1':
            print("Выберете класс " + name_star)
            print('G\nK\nM')
            inp = input()
        if inp == '-2':
            rand = -1
            inp = '-1'
        if inp == '-1':
            inp = random.randint(1, 100)
            if inp >= 97:
                return 'G'
            elif inp >= 88:
                return 'K'
            else:
                return 'M'
        else:
            return inp
    elif type_star == 4:
        if inp != '-1':
            print("Выберете класс " + name_star)
            print('Gg\nKg\nMg')
            inp = input()
        if inp == '-2':
            rand = -1
            inp = '-1'
        if inp == '-1':
            inp = random.randint(1, 100)
            if inp >= 97:
                return 'Gg'
            elif inp >= 88:
                return 'Kg'
            else:
                return 'Mg'
        else:
            return inp
    elif type_star == 5:
        if inp != '-1':
            print("Выберете класс " + name_star)
            print('Gc\nKc\nMc')
            inp = input()
        if inp == '-2':
            rand = -1
            inp = '-1'
        if inp == '-1':
            inp = random.randint(1, 100)
            if inp >= 97:
                return 'Gc'
            elif inp >= 88:
                return 'Kc'
            else:
                return 'Mc'
        else:
            return inp
    elif type_star == 6:
        if inp != '-1':
            print("Выберете класс " + name_star)
            print('L\nT\nY')
            inp = input()
        if inp == '-2':
            rand = -1
            inp = '-1'
        if inp == '-1':
            inp = random.randint(1, 13)
            if inp >= 8:
                return 'L'
            elif inp >= 2:
                return 'T'
            else:
                return 'Y'
        else:
            return inp
    elif type_star == 7:
        return 'D'
    elif type_star == 8:
        return 'Dd'
    else:
        return 'N'

def ColorStar(klass_star):
    if klass_star == 'O' or klass_star == 'Oc':
        return 'Голубой'
    elif klass_star == 'B' or klass_star == 'Bc':
        return 'Бело-голубой'
    elif klass_star == 'A' or klass_star == 'Ac' or klass_star == 'D':
        return 'Белый'
    elif klass_star == 'F' or klass_star == 'Fc':
        return 'Жёлто-белый'
    elif klass_star == 'G' or klass_star == 'Gc' or klass_star == 'Gg':
        return 'Жёлтый'
    elif klass_star == 'K' or klass_star == 'Kc' or klass_star == 'Kg':
        return 'Оранжевый'
    elif klass_star == 'M' or klass_star == 'Mc' or klass_star == 'Mg':
        return 'Красный'
    elif klass_star == 'L' or klass_star == 'T' or klass_star == 'Y':
        return 'Коричневый'
    elif klass_star == 'Dd':
        return 'Чёрный'
    elif klass_star == 'N':
        return 'Фиолетово-белый'

def TempStar(klass_star):
    if klass_star == 'O' or klass_star == 'Oc':
        return random.randint(300, 600)*100
    elif klass_star == 'B' or klass_star == 'Bc':
        return random.randint(100, 299)*100
    elif klass_star == 'A' or klass_star == 'Ac':
        return random.randint(750, 999)*10
    elif klass_star == 'F' or klass_star == 'Fc':
        return random.randint(600, 749)*10
    elif klass_star == 'G' or klass_star == 'Gc' or klass_star == 'Gg':
        return random.randint(500, 599)*10
    elif klass_star == 'K' or klass_star == 'Kc' or klass_star == 'Kg':
        return random.randint(350, 499)*10
    elif klass_star == 'M' or klass_star == 'Mc' or klass_star == 'Mg':
        return random.randint(200, 349)*10
    elif klass_star == 'L':
        return random.randint(130, 299)*10
    elif klass_star == 'T':
        return random.randint(70, 129)*10
    elif klass_star == 'Y':
        return random.randint(20, 69)*10
    elif klass_star == 'Dd':
        return random.randint(1, 100)
    elif klass_star == 'D':
        return random.randint(60, 200)*100
    elif klass_star == 'N':
        return random.randint(100, 1000)*1000

 # def RadiusStar(klass_star)

random.seed()
tmp_yn = input("Сгенерировать название системы(y/n)? ")
if tmp_yn == 'y':
    name = ""
    data_syllables = ['ал', 'ле', 'лу', 'ге', 'за', 'се', 'на', 'би', 'со', 'ус', 'юс', 'ес', 'ар', 'ма', 'ин', 'ди',
                      'ре', 'ри', 'а', 'ер', 'ат', 'ен', 'бе', 'ра', 'ла', 'ве', 'ти', 'ед', 'ор', 'ку', 'ан', 'те',
                      'ис', 'он', 'фо', 'ке', 'ек', 'ос']
    tmp_count = random.randint(1, 6)
    if tmp_count == 1:
        tmp_count = 2
    elif tmp_count == 2 or tmp_count == 3:
        tmp_count = 2
    elif tmp_count == 4 or tmp_count == 5:
        tmp_count = 3
    elif tmp_count == 6:
        tmp_count = 4
    for i in range(tmp_count):
        if random.randint(0, 1) == 0:
            name += data_syllables[random.randint(0, len(data_syllables)-1)][::-1]
        else:
            name += data_syllables[random.randint(0, len(data_syllables) - 1)]
    name = name.title()
else:
    name = input("Название системы: ")
print("Если хотите полного рандома в любой при очередном вводе напишите '-2'")
print("Если хотите локального рандома то напишите '-1'")
n_star = CountStar()
wb_template = load_workbook('template.xlsx')  # загрузка шаблона
wb = wb_template  # копирование шаблона
wsys = wb['Система']  # делаем активным основной лист
for i in range(1, n_star+1):

    if i > 25:    # литера для работы с файлом
        wlitter = chr(i // 26 + 64) + chr(i - (i // 26) * 26 + 65)
    else:
        wlitter = chr(i + 65)
    if i-1 > 25:  # литера для названия
        litter = chr((i-1) // 26 + 64) + chr((i-1) - ((i-1) // 26) * 26 + 65)
    else:
        litter = chr((i-1) + 65)
    if n_star == 1:
        name_star = name
    else:
        name_star = name+' '+litter

    n_planet = CountPlanet(name_star)
    type_star = TypeStar(name_star)
    klass_star = KlassStar(name_star, type_star)
    color_star = ColorStar(klass_star)
    temp_star = TempStar(klass_star)

    wsys[wlitter + '1'] = name_star
    wsys[wlitter + '2'] = color_star
    wsys[wlitter + '3'] = TypesStar[type_star-1]
    wsys[wlitter + '4'] = klass_star
    wsys[wlitter + '5'] = n_planet
    wsys[wlitter + '6'] = temp_star
    if n_planet > 0:
        wstar = wb.copy_worksheet(wb_template['Звезда'])
        wstar.title = name_star
        wstar['B1'] = name_star
        wstar['B2'] = color_star
        wstar['B3'] = TypesStar[type_star - 1]
        wstar['B4'] = klass_star
        wstar['B5'] = n_planet
        wstar['B6'] = temp_star



# wb.move_sheet(ws, offset=-1)
del wb["Звезда"]
del wb["Планета"]
wb.worksheets[0].name = "Main"
wb.save(name+'.xlsx')