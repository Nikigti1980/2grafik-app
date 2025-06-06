import pandas as pd
import random
from calendar import monthrange
from datetime import datetime, timedelta
import streamlit as st
from io import BytesIO

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

st.title("Генератор на работен график")

# Базов коефициент на натовареност на обекта
base_complexity = st.number_input("Въведете коефициент на сложност на обекта (пример 1.8)", min_value=1.0, step=0.1, value=1.0)

st.subheader("Настройка на работното време и пик часове по дни от седмицата")
working_hours = {}
peak_hours = {}
day_load_factors = {}
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

for day in weekdays:
    st.markdown(f"### {day}")
    col1, col2, col3 = st.columns(3)
    with col1:
        start_time = st.time_input(f"Начало на работа ({day})", value=datetime.strptime("09:00", "%H:%M").time(), key=f"start_{day}")
        end_time = st.time_input(f"Край на работа ({day})", value=datetime.strptime("21:00", "%H:%M").time(), key=f"end_{day}")
    with col2:
        peak_start_time = st.time_input(f"Начало на пик ({day})", value=datetime.strptime("14:00", "%H:%M").time(), key=f"peak_start_{day}")
        peak_end_time = st.time_input(f"Край на пик ({day})", value=datetime.strptime("18:00", "%H:%M").time(), key=f"peak_end_{day}")
    with col3:
        load_factor = st.slider(f"Натовареност (%) ({day})", min_value=0, max_value=100, value=30, key=f"load_{day}")

    working_hours[day] = (start_time, end_time)
    peak_hours[day] = (peak_start_time, peak_end_time)
    day_load_factors[day] = load_factor / 100  # превръщаме го в коефициент

uploaded_file = st.file_uploader("Качи Excel файл с таб 'Обобщение'", type=["xlsx"])

year = st.number_input("Въведете година:", min_value=2000, max_value=2100, value=datetime.now().year)
month = st.number_input("Въведете месец (1-12):", min_value=1, max_value=12, value=datetime.now().month)

if uploaded_file is not None:
    if st.button("Генерирай график"):
        try:
            summary_df = pd.read_excel(uploaded_file, sheet_name='Обобщение')
            employees_plan = summary_df[['Име', 'Планирани работни часове']]

            num_days = monthrange(year, month)[1]
            days = [datetime(year, month, day) for day in range(1, num_days + 1)]

            shift_durations = [8, 6, 4]  # часове

            schedule = []

            employee_pool = employees_plan.set_index('Име').to_dict()['Планирани работни часове']
            employees_list = list(employee_pool.keys())
            random.shuffle(employees_list)
            employee_index = 0

            for day in days:
                weekday_name = day.strftime('%A')

                start_time, end_time = working_hours[weekday_name]
                total_open_hours = (datetime.combine(day, end_time) - datetime.combine(day, start_time)).seconds // 3600
                total_needed_hours = int(total_open_hours * base_complexity * (1 + day_load_factors[weekday_name]))

                # Планиране на смените
                current_time = datetime.combine(day, start_time)
                shifts = []
                while total_needed_hours > 0:
                    for duration in shift_durations:
                        if total_needed_hours >= duration:
                            shifts.append(duration)
                            total_needed_hours -= duration
                            break
                    else:
                        if total_needed_hours > 0:
                            shifts.append(4)
                            total_needed_hours = 0

                # Разпределение на смените със застъпване
                shift_times = []
                for i, duration in enumerate(shifts):
                    start_shift = current_time
                    end_shift = start_shift + timedelta(hours=duration)
                    shift_times.append((start_shift, end_shift, duration))
                    if i % 2 == 0:
                        current_time = end_shift - timedelta(hours=2)
                    else:
                        current_time = end_shift - timedelta(hours=1)

                # Назначаване на служители с ротация
                for shift_start, shift_end, hours in shift_times:
                    employee = employees_list[employee_index % len(employees_list)]
                    employee_index += 1

                    if employee_pool[employee] >= hours:
                        employee_pool[employee] -= hours
                        schedule.append({
                            'Дата': day.strftime('%Y-%m-%d'),
                            'Ден': weekday_name,
                            'Служител': employee,
                            'Начало': shift_start.strftime('%H:%M'),
                            'Край': shift_end.strftime('%H:%M'),
                            'Продължителност (часове)': hours
                        })

            schedule_df = pd.DataFrame(schedule)

            summary_report = schedule_df.groupby('Служител')['Продължителност (часове)'].sum().reset_index()
            summary_report = summary_report.merge(employees_plan, left_on='Служител', right_on='Име')
            summary_report['Статус'] = summary_report.apply(lambda row: 'ОК' if row['Продължителност (часове)'] <= row['Планирани работни часове'] else 'НАД', axis=1)

            st.success("Графикът беше успешно генериран!")

            st.download_button(
                label="Изтегли график (Excel)",
                data=to_excel(schedule_df),
                file_name='grafik_magazin.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            st.download_button(
                label="Изтегли отчет (Excel)",
                data=to_excel(summary_report),
                file_name='grafik_summary.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            st.error(f"Възникна грешка при обработката на файла: {e}")
