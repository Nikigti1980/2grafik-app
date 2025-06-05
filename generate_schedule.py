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
st.write("Streamlit е стартиран успешно.")

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
                '8h': 9  # 8 работни + 1 присъствен час за почивка
            }

            work_patterns = [
                (3, 2),
                (3, 1),
                (4, 2),
                (2, 1)
            ]

            schedule = []

            shift_start_times = ['08:00', '09:00', '10:00', '11:00', '12:00']

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

                        remaining = planned_hours - total_hours

                        if day.weekday() in [4, 5]:  # Петък или събота
                            shift = random.choices(['8h', '6h'], weights=[0.7, 0.3])[0]
                        else:
                            shift = random.choices(['8h', '6h', '4h'], weights=[0.5, 0.3, 0.2])[0]

                        hours = shift_types[shift]

                        start_time = random.choice(shift_start_times)
                        start_dt = datetime.strptime(start_time, '%H:%M')
                        end_dt = (start_dt + timedelta(hours=(hours if shift != '8h' else 8))).time()

                        schedule.append({
                            'Дата': day.strftime('%Y-%m-%d'),
                            'Ден': day.strftime('%A'),
                            'Служител': name,
                            'Начало': start_time,
                            'Край': end_dt.strftime('%H:%M'),
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
