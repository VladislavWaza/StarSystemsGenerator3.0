from openpyxl import Workbook, load_workbook
import random
import math
rand = 0
RadiusSun = 6.9551 * 10 ** 8
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
        return random.choice(list([8]*2+[7]*4+[6]*5+[5]*10+[4]*20+[3]*60+[2]*100+[1]*799))
    else:
        return inp

def TypeStar(name_star):
    global TypesStar, rand
    if rand == -1:
        inp = -1
    else:
        print("Выберете тип " + name_star)
        for _ in range(1, 11):
            print(str(_) + '.' + TypesStar[_ - 1])
        inp = int(input())
        if inp == -2:
            rand = -1
            inp = -1
    if inp == -1:
        return random.choice(list([1]*5+[2]*1+[3]*74+[4]*4+[5]*2+[6]*5+[7]*5+[8]*1+[9]*2+[10]))
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
        return random.choice(list([0,1,2,3,3,4,4,5,5,5,6,6,6,7,7,7,8,8,9,9,10,11,12,13,14,15,16,17,18,19,20]))
    else:
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
            return random.choice(list(['F']*71+['A']*23+['B']*4+['O']*2))
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
            return random.choice(list(['Fc']*71+['Ac']*23+['Bc']*4+['Oc']*2))
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
            return random.choice(list(['G']*4+['K']*9+['M']*87))
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
            return random.choice(list(['Gg'] * 4 + ['Kg'] * 9 + ['Mg'] * 87))
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
            return random.choice(list(['Gc'] * 4 + ['Kc'] * 9 + ['Mc'] * 87))
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
            return random.choice(list(['L']*6+['T']*6+['Y']))
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

def RadiusStar(klass_star):
    if klass_star == 'O':
        inp = random.randint(105, 195)/10
    elif klass_star == 'B':
        inp = random.randint(49, 91)/10
    elif klass_star == 'A':
        inp = random.randint(15, 27)/10
    elif klass_star == 'F':
        inp = random.randint(9, 17)/10
    elif klass_star == 'G':
        inp = random.randint(77, 140)/100
    elif klass_star == 'K':
        inp = random.randint(60, 110)/100
    elif klass_star == 'M':
        inp = random.randint(9, 71)/100
    elif klass_star == 'L' or klass_star == 'T' or klass_star == 'Y':
        inp = random.randint(12, 70)/1000
    elif klass_star == 'Dd' or klass_star == 'D':
        return '-', random.randint(609, 819)*1000
    elif klass_star == 'N':
        return '-', random.randint(100, 150)*100
    elif len(klass_star) == 2 and klass_star[1] == 'g':
        inp = random.randint(10, 100)
    elif len(klass_star) == 2 and klass_star[1] == 'c':
        inp = random.randint(10, 180)*10
    return inp, inp * RadiusSun


random.seed()
tmp_yn = input("Сгенерировать название системы(y/n)? ")
if tmp_yn == 'y':
    name = ""
    data_syllables = ['ал', 'ле', 'лу', 'ге', 'за', 'се', 'на', 'би', 'со', 'ус', 'юс', 'ес', 'ар', 'ма', 'ин', 'ди',
                      'ре', 'ри', 'а', 'ер', 'ат', 'ен', 'бе', 'ра', 'ла', 'ве', 'ти', 'ед', 'ор', 'ку', 'ан', 'те',
                      'ис', 'он', 'фо', 'ке', 'ек', 'ос']
    tmp_count = random.choice([2,3,3,4,4,5])
    for i in range(tmp_count):
        if random.randint(0, 1) == 0:
            name += random.choice(data_syllables)[::-1]
        else:
            name += random.choice(data_syllables)
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
    rad_star_sun, rad_star_m = RadiusStar(klass_star)

    wsys[wlitter + '1'] = name_star
    wsys[wlitter + '2'] = color_star
    wsys[wlitter + '3'] = TypesStar[type_star-1]
    wsys[wlitter + '4'] = klass_star
    wsys[wlitter + '5'] = n_planet
    wsys[wlitter + '6'] = temp_star
    wsys[wlitter + '7'] = rad_star_m
    wsys[wlitter + '8'] = rad_star_sun
    if n_planet > 0:
        wstar = wb.copy_worksheet(wb_template['Звезда'])
        wstar.title = name_star
        wstar['B1'] = name_star
        wstar['B2'] = color_star
        wstar['B3'] = TypesStar[type_star - 1]
        wstar['B4'] = klass_star
        wstar['B5'] = n_planet
        wstar['B6'] = temp_star
        wstar['B7'] = rad_star_m
        wstar['B8'] = rad_star_sun



# wb.move_sheet(ws, offset=-1)
del wb["Звезда"]
del wb["Планета"]
wb.worksheets[0].name = "Main"
wb.save(name+'.xlsx')