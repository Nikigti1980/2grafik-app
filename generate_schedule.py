import pandas as pd
import random
from calendar import monthrange
from datetime import datetime, timedelta
import os

# Зареждане на файла с данни
file_path = 'input_schedule.xlsx'  # <-- тук качваш твоя файл
summary_df = pd.read_excel(file_path, sheet_name='Обобщение')

# Взимане на нужните данни
employees_plan = summary_df[['Име', 'Планирани работни часове']]

# Въвеждане на работен месец и година
year = int(input("Въведете година (напр. 2025): "))
month = int(input("Въведете месец (1-12): "))

num_days = monthrange(year, month)[1]
days = [datetime(year, month, day) for day in range(1, num_days + 1)]

# Дефиниране на типа смени
shift_types = {
    '4h': 4,
    '6h': 6,
    '8h': 9  # 8 работни + 1 присъствен час за почивка
}

# Силни дни (повече персонал)
strong_days = [4, 5]  # Петък=4, Събота=5

# Логика за работни и почивни дни
work_patterns = [
    (3, 2),  # оптимално
    (3, 1),  # динамично
    (4, 2),  # тежко натоварване
    (2, 1)   # леко натоварване
]

# Инициализиране на график
schedule = []

# Стартиране на планирането за всеки служител
for idx, row in employees_plan.iterrows():
    name = row['Име']
    planned_hours = row['Планирани работни часове']

    day_pointer = 0
    total_hours = 0
    pattern_idx = 0

    while day_pointer < len(days) and total_hours < planned_hours:
        work_days, rest_days = work_patterns[pattern_idx % len(work_patterns)]
        # Работни дни
        for _ in range(work_days):
            if day_pointer >= len(days) or total_hours >= planned_hours:
                break

            day = days[day_pointer]

            # Избор на смяна, като предпочитаме 8h, но ако наближаваме лимита - 4h или 6h
            remaining = planned_hours - total_hours
            if remaining >= 8:
                shift = '8h'
            elif remaining >= 6:
                shift = '6h'
            else:
                shift = '4h'

            hours = shift_types[shift]

            schedule.append({
                'Дата': day.strftime('%Y-%m-%d'),
                'Ден': day.strftime('%A'),
                'Служител': name,
                'Смяна': shift,
                'Часове': hours
            })

            total_hours += (hours if shift != '8h' else 8)  # Само 8ч работа, 9ч присъствие
            day_pointer += 1

        # Почивни дни
        day_pointer += rest_days
        pattern_idx += 1

# Преобразуване в DataFrame
schedule_df = pd.DataFrame(schedule)

# Проверка на реално планираните часове
summary_report = schedule_df.groupby('Служител')['Часове'].sum().reset_index()
summary_report = summary_report.merge(employees_plan, left_on='Служител', right_on='Име')
summary_report['Статус'] = summary_report.apply(lambda row: 'ОК' if row['Часове'] <= row['Планирани работни часове'] else 'НАД', axis=1)

# Запазване в директорията Documents
save_directory = os.path.expanduser('~/Documents/Grafici')
os.makedirs(save_directory, exist_ok=True)

base_filename = os.path.join(save_directory, 'grafik_magazin.xlsx')
summary_filename = os.path.join(save_directory, 'grafik_summary.xlsx')
filename = base_filename

counter = 1
while True:
    try:
        schedule_df.to_excel(filename, index=False)
        summary_report.to_excel(summary_filename, index=False)
        break  # Успешно записани файлове
    except PermissionError:
        filename = os.path.join(save_directory, f'grafik_magazin_{counter}.xlsx')
        summary_filename = os.path.join(save_directory, f'grafik_summary_{counter}.xlsx')
        counter += 1

print(f"Графикът е успешно генериран и записан в '{filename}'")
print(f"Отчетът е успешно записан в '{summary_filename}'")
