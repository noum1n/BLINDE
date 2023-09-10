import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mplcursors  # Импортируем mplcursors для подсказок при наведении

# Загрузите данные из Excel файла
file_path = 'optimal_routes.xlsx'
df = pd.read_excel(file_path)

# Преобразование дат в формат datetime
df['Дата и время начала'] = pd.to_datetime(df['Дата и время начала'], format='%d.%m.%y %H:%M')
df['Дата и время окончания'] = pd.to_datetime(df['Дата и время окончания'], format='%d.%m.%y %H:%M')

# Создание графика
fig, ax = plt.subplots(figsize=(12, 8))

# Установите цвета для судов и ледоколов
colors = {' - ': 'black', 'Таймыр': 'blue', 'Ямал': 'red', '50 лет Победы': 'blue', 'Вайгач': 'black'}

# Создайте словарь для сопоставления судов с цветами
ship_colors = {ship: colors[icebreaker] for ship, icebreaker in zip(df['Судно'], df['Ледокол'])}

# Сортировка данных по времени начала
df.sort_values(by='Дата и время начала', inplace=True)

# Постройте графики для каждого судна
for index, row in df.iterrows():
    start_date = row['Дата и время начала']
    end_date = row['Дата и время окончания']
    ship = row['Судно']
    color = ship_colors[ship]

    # Отобразите данные на графике
    ax.hlines(y=ship, xmin=start_date, xmax=end_date, color=color, linewidth=4, label=ship)

# Настройте внешний вид графика
plt.xlabel('Время')
plt.ylabel('Судно')
plt.title('График движения судов с учетом ледоколов')

# Улучшение подписей шкалы времени
ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%y"))
ax.xaxis.set_minor_locator(mdates.HourLocator(interval=12))

# Установка поворота подписей на шкале времени
plt.xticks(rotation=45)

# Увеличение размера шрифта
plt.xticks(fontsize=10)
plt.yticks(fontsize=12)

# Выведите график
plt.grid(axis='x')
plt.tight_layout()

# Добавление всплывающих подсказок при наведении
mplcursors.cursor(hover=True).connect("add", lambda sel: sel.annotation.set_text(f'{sel.artist.get_label()}\n'
                                                                             f'{sel.target.get_xdata()[0].strftime("%d.%m.%y %H:%M")} - '
                                                                             f'{sel.target.get_xdata()[1].strftime("%d.%m.%y %H:%M")}'))

plt.show()
