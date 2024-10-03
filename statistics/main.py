import pandas as pd
import os
import keyboard
from math import sqrt
import matplotlib.pyplot as plt


def fact(x: int):
    res = 1
    for i in range(1, x + 1):
        res *= i
    return res

def expection(data: dict):
    return round(sum([i * data[i] for i in data.keys()]), 2)


#main data
tries = int(input('Введите количество попыток:\n'))
success = float(input('Введите вероятность (в виде дроби):\n'))
while success < 0 or success > 1:
    success = float(input('Такой вероятности не существует. Попробуйте еще раз:\n'))
fail = 1 - success
results = dict()
for attempt in range(tries + 1):
    results[attempt] = round((fact(tries) / (fact(attempt) * fact(tries - attempt))) * success**attempt * fail**(tries - attempt), 4)
results_percent = dict(zip([i for i in range(tries + 1)], [f'{round(n * 100, 2)} %' for n in results.values()]))

#table
df = pd.DataFrame({'Количество успешных испытаний': results_percent.keys(), 'Вероятность, %': results_percent.values()})
df.to_excel('distribution_table.xlsx', index=False)

#text
matexpection = expection(results)
data_for_dispersion = dict(zip([round((i - matexpection)**2, 2) for i in results.keys()], results.values()))
dispersion = expection(data_for_dispersion)
with open('statistic_values.txt', 'w') as f:
    f.write(f'Математическое ожидание успешных попыток: {matexpection}\n\n')
    if tries <= 10:
        f.write(f'Расчёты: {" + ".join([f"{i} * {results[i]}" for i in results.keys()])}')
    else:
        f.write('Слишком большое количество данных для отображения расчётов')
    f.write(f'\n\nДисперсия: {dispersion}\n\nСтандартное отклонение: {round(sqrt(dispersion), 2)}')
f.close()

#plot
plt.plot(results.keys(), results.values())
plt.savefig('diagram.jpeg')

#opening files
os.startfile('distribution_table.xlsx')
os.startfile('statistic_values.txt')
os.startfile('diagram.jpeg')

#closing files
print('\nНажмите ПРОБЕЛ для закрытия файлов')
keyboard.wait('Space')
os.system('taskkill /IM EXCEL.EXE')
os.system('taskkill /IM notepad.exe')
