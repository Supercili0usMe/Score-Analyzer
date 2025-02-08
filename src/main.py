import data_processing as d
import visualization as v
from openpyxl import load_workbook
from collections import defaultdict

def getWorksheet(fileName):
    filePath = d.get_file_path(fileName, d.folder_root)
    data = load_workbook(filePath)
    worksheet = data.active
    return worksheet

def printInfo(worksheet, subjects, allMarks):
    means = []
    info = d.extract_info(worksheet)
    print('\n')
    for i in info:
        print(f'{i}: {info[i]}')
    print('\nСредний балл по всем предметам:\n')

    for i in subjects:
        subject = subjects[i]
        period = info['Период']
        estInf = v.countMean(subject, allMarks, period)
        try:
            means.append(float(estInf[estInf.index('-') + 1:estInf.index('~')]))
        except:
            pass
        print(estInf)
    totalScore = v.countTotalScore(allMarks)
    print(f'___________________________\nСреднее всех средних баллов - {v.setColorOfScore(totalScore)}{totalScore}{'\033[0m'}')
    print(f'Итого: {v.totalEst[0]} пятёрок; {v.totalEst[1]} четверок; {v.totalEst[2]} троек; {v.totalEst[3]} двоек; {v.totalEst[4]} единиц; не хватает оценок у {len(subjects) - sum(v.totalEst)} предметов')

def process_grades(grades, dates):
    # Группируем оценки по датам
    date_to_grades = defaultdict(list)
    for date, grade in zip(dates, grades):
        date_to_grades[date].append(grade)

    # Сохраняем уникальные даты в порядке их первого появления
    unique_dates = []
    seen_dates = set()
    for date in dates:
        if date not in seen_dates:
            unique_dates.append(date)
            seen_dates.add(date)

    # Вычисляем средние оценки для каждой даты
    avg_grades = []
    for date in unique_dates:
        grades_list = date_to_grades[date]
        average = sum(grades_list) / len(grades_list)
        avg_grades.append(average)

    return avg_grades, unique_dates

def drawGraph(subForGraph, subjects, allMarks):
    if subForGraph in subjects.values():
        scores = v.extractScoreMass(subForGraph, allMarks)
        dates = d.refactor_marks(allMarks, subForGraph)[0]
        scores, dates = process_grades(scores, dates)
        if len(scores) <= 1:
            print(f'{'\033[31m'}Слишком мало оценок для отрисовки графика{'\033[0m'}')
            exit()
        else:
            v.drawGraph(subForGraph, scores, dates)
    elif subForGraph != '':
        print(f'{'\033[31m'}Указан несуществующий предмет{'\033[0m'}')


def main():
    print(f">_\nДля того, чтобы прога работала, нужно создать папку data в корне проекта, затем закинуть туда файл с оценками с расширением .xlsx и перезапустить программу\n"
        "Эта прога создана для тестинга на работоспособность, поэтому скоро будет готова telegram bot версия")
    fileName = input("Введите имя файла (например, example.xlsx): ").strip()
    if not '.xlsx' in fileName:
        print(f'{'\033[31m'}Файл нечитаем, должно быть расширение{'\033[1;91m'} .xlsx{'\033[0m'}')
        return
    try:
        worksheet = getWorksheet(fileName)
        subjects = d.extract_subjects(worksheet)
        allMarks = d.extract_marks(worksheet, subjects)
        if allMarks == "":
            print("\033[1;91mВ вашем файле отсутствуют комментарии к отметкам, их наличие критически важно\033[0m")
            return
        printInfo(worksheet, subjects, allMarks)
        subForGraph = input('\nГрафик изменения среднего балла какого предмета нарисовать (если не надо рисовать, то нажмите enter) ')
        drawGraph(subForGraph, subjects, allMarks)
    except TypeError:
        print("\033[31m\033[1;91mПроизошла непредвиденная ошибка, пожалуйста повторите запрос\033[0m")


if __name__ == '__main__':
    main()
