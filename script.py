from openpyxl import load_workbook

# указываем названия файлов и загружаем их
for file in range(0, 11):
    fn = f'{file}.xlsx'
    wb = load_workbook(fn, data_only=True)

    # делаем первый лист активным
    wb.active = 0
    sheet = wb.active

    # получаем сумму тишины, затем значения операторов
    silence_sum = 0

    for D in range(2, 12):
        value = sheet[f'D{D}'].value
        if value is None:
            value = 0
        silence_sum = silence_sum + value

    silence_sum_tele2 = sheet['J2'].value + sheet['J3'].value
    silence_sum_mts = sheet['J6'].value + sheet['J7'].value
    silence_sum_megafon = sheet['J8'].value + sheet['J9'].value
    silence_sum_beeline = 0
    try:
        silence_sum_beeline = sheet['J10'].value + sheet['J11'].value
    except Exception as error:
        print(f'Ошибка {error}, вероятнее всего нет данных!')

    # получаем процентаж тишины для операторов
    silence_percent_tele2 = round(silence_sum_tele2*100 / silence_sum, 2)
    silence_percent_mts = round(silence_sum_mts*100 / silence_sum, 2)
    silence_percent_megafon = round(silence_sum_megafon*100 / silence_sum, 2)
    if silence_sum_beeline == 0:
        print('Нету оператора билайн!')
    else:
        silence_percent_beeline = round(silence_sum_beeline*100 / silence_sum, 2)

    # красиво вносим полученные расчеты
    sheet['A20'] = 'Tele2'
    sheet['B20'] = 'MTS'
    sheet['C20'] = 'Megafon'
    sheet['D20'] = 'Beeline'

    sheet['A21'] = silence_percent_tele2
    sheet['B21'] = silence_percent_mts
    sheet['C21'] = silence_percent_megafon
    sheet['D21'] = silence_sum_beeline

    # не забываем сохранить файл
    wb.save(fn)
