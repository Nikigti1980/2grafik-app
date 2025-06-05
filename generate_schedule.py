import pandas as pd
import random
from calendar import monthrange
from datetime import datetime, timedelta
import os
import streamlit as st
from io import BytesIO

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

st.title("Генератор на работен график")

st.subheader("Настройка на работното време по дни от седмицата")
working_hours = {}
weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

col1, col2 = st.columns(2)
with col1:
    default_start_time = st.time_input("Начало на работа (за всички дни)", value=datetime.strptime("09:00", "%H:%M").time(), key="start_default")
with col2:
    default_end_time = st.time_input("Край на работа (за всички дни)", value=datetime.strptime("21:00", "%H:%M").time(), key="end_default")

st.markdown("**При нужда коригирай работното време за конкретен ден:**")

for day in weekdays:
    col1, col2 = st.columns(2)
    with col1:
        start_time = st.time_input(f"Начало на работа ({day})", value=default_start_time, key=f"start_{day}")
    with col2:
        end_time = st.time_input(f"Край на работа ({day})", value=default_end_time, key=f"end_{day}")
    working_hours[day] = (start_time, end_time)

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

            shift_types = {
                '4h': 4,
                '6h': 6,
                '8h': 9
            }

            work_patterns = [
                (3, 2),
                (3, 1),
                (4, 2),
                (2, 1)
            ]

            schedule = []

            for idx, row in employees_plan.iterrows():
                name = row['Име']
                planned_hours = row['Планирани работни часове']

                day_pointer = idx % 7
                total_hours = 0
                pattern_idx = idx % len(work_patterns)

                while day_pointer < len(days) and total_hours < planned_hours:
                    work_days, rest_days = work_patterns[pattern_idx]

                    for _ in range(work_days):
                        if day_pointer >= len(days) or total_hours >= planned_hours:
                            break

                        day = days[day_pointer]
                        weekday_name = day.strftime('%A')
                        start_time, end_time = working_hours[weekday_name]

                        total_work_hours = (datetime.combine(day, end_time) - datetime.combine(day, start_time)).seconds // 3600

                        remaining = planned_hours - total_hours

                        if total_work_hours >= 8 and remaining >= 8:
                            shift = '8h'
                        elif total_work_hours >= 6 and remaining >= 6:
                            shift = '6h'
                        else:
                            shift = '4h'

                        hours = shift_types[shift]

                        shift_start_hour = start_time.hour
                        shift_start_minute = start_time.minute
                        shift_start_dt = datetime(year, month, day.day, shift_start_hour, shift_start_minute)
                        shift_end_dt = (shift_start_dt + timedelta(hours=(hours if shift != '8h' else 8))).time()

                        schedule.append({
                            'Дата': day.strftime('%Y-%m-%d'),
                            'Ден': weekday_name,
                            'Служител': name,
                            'Начало': start_time.strftime('%H:%M'),
                            'Край': shift_end_dt.strftime('%H:%M'),
                            'Смяна': shift,
                            'Часове': hours
                        })

                        total_hours += (hours if shift != '8h' else 8)
                        day_pointer += 1

                    day_pointer += rest_days
                    pattern_idx = (pattern_idx + 1) % len(work_patterns)

            schedule_df = pd.DataFrame(schedule)

            summary_report = schedule_df.groupby('Служител')['Часове'].sum().reset_index()
            summary_report = summary_report.merge(employees_plan, left_on='Служител', right_on='Име')
            summary_report['Статус'] = summary_report.apply(lambda row: 'ОК' if row['Часове'] <= row['Планирани работни часове'] else 'НАД', axis=1)

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
