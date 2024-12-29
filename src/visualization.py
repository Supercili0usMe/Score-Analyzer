import data_processing as d

def printMean(subject, marks, coeffs):
    """
    Считаем средний балл

    Принимает:
        subject - название предмета
        marks - массив оценок для данного предмета
        coeffs - массив коэффицентов для данного предмета

    Возвращает для вывода в цикле:
        {subject} - {score} ~ {roundScore}, где subject - предмет; score - средний балл;
         roundScore - округленный средний балл
    """
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
        return f'{subject} - {score} ~ {color}{roundScore}{'\033[0m'}'

def extractScoreMass(subject, allMarks):
    """
    Создаем массив изменений среднего балла

    Принимает:
        subject - название предмета (если нужно несколько предметов то предметы через пробел)
        allMarks - массив всех оценок для каждого предмета
        coeffs - массив коэффициентов
    Возвращает:
        простой массив изменений среднего балла
    """
    marks = []
    coeffs = []
    scores = []
    if ' ' in subject: # если предметов несколько
        subject = subject.split()
        for i in subject:
            marks += d.refactor_marks(allMarks, i)[1]
            coeffs += d.refactor_marks(allMarks, i)[2]
    else:
        marks = d.refactor_marks(allMarks, subject)[1]
        coeffs = d.refactor_marks(allMarks, subject)[2]
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
    if ' ' in subject:
        subject = subject.split()
        plt.title('График изменения среднего балла по выбранным предметам')
        dates = set(dates) # убрать повторяющиеся даты
        dates = [dt.strptime(i, '%d.%m') for i in dates]
        dates.sort() # сотрировка дат
        dates = [dt.strftime(i, '%d.%m') for i in dates]
    else:
        plt.title(f'График изменения среднего балла по предмету {subject}')
        plt.ylim((math.ceil(min(scores) - 0.9) if math.ceil(min(scores) - 1) >= 2 else 2) - 0.07,
                 (math.floor(max(scores) + 0.9) if math.floor(max(scores) + 1) <= 5 else 5) + 0.07)
        plt.plot(dates, scores, 'r-o')
        plt.grid()
        plt.show()