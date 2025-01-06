import data_processing as d
from openpyxl import load_workbook
import os

totalEst = [0, 0, 0, 0]

def countMean(subject, allMarks):
    """
    Считаем средний балл

    Принимает:
        subject - название предмета (если нужно несколько предметов то предметы через пробел)
        allMarks - массив всех оценок для каждого предмета

    Возвращает для вывода в цикле:
        {subject} - {score} ~ {roundScore}, где subject - предмет; score - средний балл;
         roundScore - округленный средний балл
    """
    marks = d.refactor_marks(allMarks, subject)[1]
    coeffs = d.refactor_marks(allMarks, subject)[2]
    if marks == [] or coeffs == []:
        return f'{subject} - нет оценок'
    minNumOfEstimates = 3
    for i in range(len(marks)):
        marks[i] = marks[i] * coeffs[i]
    score = round(sum(marks) / sum(coeffs), 2)
    roundScore = round(score)
    global color
    if roundScore == 3 or roundScore == 2: # определение цвета вывода округленного балла
        color = '\033[1;91m'  # красный
    if roundScore == 4:
        color = '\033[1;93m'  # желтый
    if roundScore == 5:
        color = '\033[1;92m'  # зеленый
    if len(marks) < minNumOfEstimates:
        return f'{subject} - {score} ~ {roundScore} (не хватает {minNumOfEstimates - len(marks)} оценок)'
    else:
        totalEst[-roundScore + 5] += 1
        return f'{subject} - {score} ~ {color}{roundScore}{'\033[0m'}'

def extractScoreMass(subject, allMarks):
    """
    Создаем массив изменений среднего балла

    Принимает:
        subject - название предмета (если нужно несколько предметов то предметы через пробел)
        allMarks - массив всех оценок для каждого предмета

    Возвращает:
        простой массив изменений среднего балла
    """
    # if ' ' in subject: # если предметов несколько
    #     subject = subject.split()
    #     for i in subject:
    #         marks += d.refactor_marks(allMarks, i)[1]
    #         coeffs += d.refactor_marks(allMarks, i)[2]
    marks = d.refactor_marks(allMarks, subject)[1]
    coeffs = d.refactor_marks(allMarks, subject)[2]
    marks = [x * y for x, y in zip(marks, coeffs)]
    scores = []
    for i in range(len(marks)):
        scores.append(round(sum(marks[:i + 1]) / sum(coeffs[:i + 1]), 2))
    return scores

def drawGraph(subject: str, scores: list, dates: list):
    """
    Рисуем график изменения среднего балла

    Принимает:
        subject - название предмета (если нужно нарисовать несколько предметов то предметы через пробел)
        scores - массив изменений среднего балла у данного предмета
        dates - массив дат оценок у данного предмета
    """
    from datetime import datetime as dt
    from matplotlib import pyplot as plt
    import math
    # if ' ' in subject:
    #     subject = subject.split()
    #     plt.title('График изменения среднего балла по выбранным предметам')
    #     dates = set(dates) # убрать повторяющиеся даты
    #     dates = [dt.strptime(i, '%d.%m') for i in dates]
    #     dates.sort() # сотрировка дат
    #     dates = [dt.strftime(i, '%d.%m') for i in dates]
    dates = [dt.strptime(i, '%d.%m.%Y') for i in dates]
    dates = [f'{str(i.day).zfill(2)}.{str(i.month).zfill(2)}' for i in dates]
    plt.title(f'График изменения среднего балла по предмету {subject}')
    # plt.ylim((math.ceil(min(scores) - 0.9) if math.ceil(min(scores) - 1) >= 2 else 2) - 0.07,
    #          (math.floor(max(scores) + 0.9) if math.floor(max(scores) + 1) <= 5 else 5) + 0.07)

    minLim = (min(scores) - 0.5 if min(scores) - 0.5 >= 2 else 2) - 0.07
    maxLim = (max(scores) + 0.5 if max(scores) + 0.5 <= 5 else 5) + 0.07
    plt.ylim(minLim, maxLim)
    colors = ['red', 'orange', 'green']
    # отрисовка линий границ изменения ср. балла (2.5, 3.5, 4.5)
    for i in [2.5, 3.5, 4.5]:
        if minLim <= i <= maxLim:
            plt.axhline(y=i, color=colors[[2.5, 3.5, 4.5].index(i)], linestyle='--')

    if len(dates) >= 35:
        plt.xticks(fontsize=8)
    plt.plot(dates, scores, 'r-o')
    plt.grid()
    plt.show()

def main():
    # получаем worksheet
    fileName = input("Введите имя файла (например, example.xlsx): ").strip()
    print()
    filePath = d.get_file_path(fileName, d.folder_root)
    data = load_workbook(filePath)
    worksheet = data.active

    subjects = d.extract_subjects(worksheet)
    allMarks = d.extract_marks(worksheet, subjects)

    # вывод всех средних баллов
    # print(d.extract_info(worksheet))
    info = d.extract_info(worksheet)
    for i in info:
        print(f'{i}: {info[i]}')
    print('\nСредний балл по всем предметам:\n')
    for i in subjects:
        subject = subjects[i]
        print(countMean(subject, allMarks))
    print()
    print(f'Итого: {totalEst[0]} пятёрок; {totalEst[1]} четверок; {totalEst[2]} троек; {totalEst[3]} двоек; не хватает оценок у {len(subjects) - sum(totalEst)} предметов')
    subForGraph = input('График изменения среднего балла какого предмета нарисовать (если не надо рисовать, то нажмите enter) ')

    # рисование графика
    if subForGraph in subjects.values():
        scores = extractScoreMass(subForGraph, allMarks)
        dates = d.refactor_marks(allMarks, subForGraph)[0]
        if len(scores) <= 1:
            print('Слишком мало оценок для отрисовки графика')
            return
        else:
            drawGraph(subForGraph, scores, dates)
    elif subForGraph != '':
        print('Указан несуществующий предмет')


if __name__ == '__main__':
    main()
