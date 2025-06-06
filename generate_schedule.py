import pandas as pd
import random
from calendar import monthrange
from datetime import datetime, timedelta, time
import streamlit as st
from io import BytesIO

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.title("Генератор на работен график - НОВА ВЕРСИЯ")

# Въвеждане на коефициент на сложност
complexity_factor = st.number_input("Въведете коефициент на сложност на обекта:", min_value=1.0, max_value=5.0, value=1.8, step=0.1)

days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

day_settings = {}

with st.form("work_time_form"):
    st.subheader("Работно време и пик за седмицата")
    default_start = st.time_input("Начало на работа (за всички дни)", time(9, 0))
    default_end = st.time_input("Край на работа (за всички дни)", time(21, 0))

    for day in days_of_week:
        st.write(f"**{day}**")
        start = st.time_input(f"Начало на работа ({day})", default_start, key=f"start_{day}")
        end = st.time_input(f"Край на работа ({day})", default_end, key=f"end_{day}")
        peak_start = st.time_input(f"Начало на пик ({day})", time(14, 0), key=f"peak_start_{day}")
        peak_end = st.time_input(f"Край на пик ({day})", time(18, 0), key=f"peak_end_{day}")
        load_percent = st.number_input(f"Процент натоварване ({day})", min_value=0, max_value=200, value=30 if day not in ['Friday', 'Saturday', 'Sunday'] else (45 if day != 'Saturday' else 60), key=f"load_{day}")
        day_settings[day] = {
            "start": start,
            "end": end,
            "peak_start": peak_start,
            "peak_end": peak_end,
            "load_percent": load_percent
        }
    st.form_submit_button("Запази работното време и натоварване")

uploaded_file = st.file_uploader("Качи Excel файл с таб 'Обобщение'", type=["xlsx"])

if uploaded_file is not None and st.button("Генерирай график"):
    try:
        summary_df = pd.read_excel(uploaded_file, sheet_name='Обобщение')
        employees_plan = summary_df[['Име', 'Планирани работни часове']]

        shifts = []
        year = datetime.now().year
        month = datetime.now().month
        num_days = monthrange(year, month)[1]
        employees = employees_plan['Име'].tolist()
        emp_idx = 0

        for day_num in range(1, num_days + 1):
            date = datetime(year, month, day_num)
            day_name = date.strftime('%A')
            settings = day_settings[day_name]

            start_time = datetime.combine(date, settings['start'])
            end_time = datetime.combine(date, settings['end'])
            total_base_hours = (end_time - start_time).seconds // 3600
            required_hours = int(total_base_hours * complexity_factor * (1 + settings['load_percent'] / 100))

            peak_start_time = datetime.combine(date, settings['peak_start'])
            peak_end_time = datetime.combine(date, settings['peak_end'])

            current_time = start_time
            shift_options = [8, 6, 4]

            while required_hours > 0 and current_time < end_time:
                possible_durations = [h for h in shift_options if h <= (end_time - current_time).seconds // 3600]
                if not possible_durations:
                    break
                shift_hours = max(possible_durations)

                shift_end_time = current_time + timedelta(hours=shift_hours)

                # Проверка дали shift_end_time излиза извън end_time
                if shift_end_time > end_time:
                    shift_end_time = end_time
                    shift_hours = (shift_end_time - current_time).seconds // 3600
                    if shift_hours == 0:
                        break

                shifts.append({
                    'Дата': date.strftime('%Y-%m-%d'),
                    'Ден': day_name,
                    'Служител': employees[emp_idx % len(employees)],
                    'Начало': current_time.strftime('%H:%M'),
                    'Край': shift_end_time.strftime('%H:%M'),
                    'Смяна': f"{shift_hours}h",
                    'Часове': shift_hours
                })
                emp_idx += 1
                if peak_start_time <= current_time <= peak_end_time:
                    current_time += timedelta(hours=shift_hours - 2)  # Застъпване по време на пик
                else:
                    current_time += timedelta(hours=shift_hours)
                required_hours -= shift_hours

        schedule_df = pd.DataFrame(shifts)

        st.success("Графикът беше успешно генериран!")

        st.download_button(
            label="Изтегли график (Excel)",
            data=to_excel(schedule_df),
            file_name='grafik_magazin.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error(f"Възникна грешка при обработката на файла: {e}")
