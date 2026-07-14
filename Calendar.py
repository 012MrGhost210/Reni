import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, date
import os

# ============================================================
# НАСТРАИВАЕМЫЕ ПАРАМЕТРЫ
# ============================================================
CALENDAR_FILE_PATH = r"M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\Календарь.xlsx"
PORTFOLIO_FILE_PATH = r"M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\NAV.xlsx"
# ============================================================

st.set_page_config(
    page_title="📅 Coupon Calendar",
    page_icon="📅",
    layout="wide"
)

# --- Загрузка данных ---
def load_calendar(file_path):
    try:
        if not os.path.exists(file_path):
            return pd.DataFrame(), f"File not found: {file_path}"
        df = pd.read_excel(file_path, skiprows=3)
        df.columns = ['ISIN', 'NAME', 'VOLUME', 'DATE', 'NOMINAL', 'CURRENCY', 
                     'OUTSTANDING_NOMINAL', 'COUPON_RATE', 'PAYMENT', 'PAYMENT_RUB']
        df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
        df = df.dropna(subset=['DATE'])
        df['ISIN'] = df['ISIN'].astype(str).str.strip()
        df = df[df['ISIN'] != 'null']
        df = df[df['ISIN'] != '']
        return df, None
    except Exception as e:
        return pd.DataFrame(), str(e)

def load_portfolio(file_path):
    try:
        if not os.path.exists(file_path):
            return pd.DataFrame(), f"File not found: {file_path}"
        df = pd.read_excel(file_path)
        # Только эти колонки
        required = ['NAV_DATE', 'PORTFOLIO', 'MANAGEMENT_COMPANY', 'ASSET', 'ISIN']
        if all(col in df.columns for col in required):
            df['NAV_DATE'] = pd.to_datetime(df['NAV_DATE'], errors='coerce')
            df['ISIN'] = df['ISIN'].astype(str).str.strip()
            return df, None
        else:
            missing = [col for col in required if col not in df.columns]
            return pd.DataFrame(), f"Missing columns: {missing}"
    except Exception as e:
        return pd.DataFrame(), str(e)

# --- Загружаем файлы ---
calendar_df, cal_err = load_calendar(CALENDAR_FILE_PATH)
portfolio_df, port_err = load_portfolio(PORTFOLIO_FILE_PATH)

if cal_err:
    st.error(f"Calendar error: {cal_err}")
    st.stop()

if port_err:
    st.warning(f"Portfolio error: {port_err}")

st.title("📅 Coupon Payment Calendar")

# Выбор месяца
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    year = st.selectbox("Year", [2026, 2027, 2028], index=0, label_visibility="collapsed")
    month = st.selectbox("Month", range(1, 13), 
                        format_func=lambda x: calendar.month_name[x],
                        index=datetime.now().month - 1,
                        label_visibility="collapsed")

st.markdown(f"## {calendar.month_name[month]} {year}")

# Создаем lookup для портфеля по ISIN
portfolio_lookup = {}
if not portfolio_df.empty:
    for _, row in portfolio_df.iterrows():
        portfolio_lookup[row['ISIN']] = {
            'ASSET': row.get('ASSET', ''),
            'PORTFOLIO': row.get('PORTFOLIO', ''),
            'MANAGEMENT_COMPANY': row.get('MANAGEMENT_COMPANY', '')
        }

# Фильтруем события за месяц
first_day = date(year, month, 1)
last_day = date(year, month, calendar.monthrange(year, month)[1])
month_events = calendar_df[
    (calendar_df['DATE'].dt.date >= first_day) & 
    (calendar_df['DATE'].dt.date <= last_day)
]

# Цвета для УК
uk_colors = {
    'ТКБ ИНВЕСТМЕНТ ПАРТНЕРС (АО)': '#FF6B6B',
    'ООО СК РЕНЕССАНС ЖИЗНЬ': '#4ECDC4',
    'УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ АО': '#45B7D1',
    'ПЕРВАЯ АО УК': '#96CEB4',
    'УК РАЙФФАЙЗЕН ООО': '#FFD93D',
}
default_color = '#D3D3D3'

# Группируем события по дням
events_by_day = {}
for _, event in month_events.iterrows():
    day = event['DATE'].day
    if day not in events_by_day:
        events_by_day[day] = []
    
    isin = event['ISIN']
    event_info = {
        'ISIN': isin,
        'PAYMENT': event.get('PAYMENT_RUB', 0),
    }
    
    if isin in portfolio_lookup:
        event_info['ASSET'] = portfolio_lookup[isin]['ASSET']
        event_info['PORTFOLIO'] = portfolio_lookup[isin]['PORTFOLIO']
        event_info['MANAGEMENT_COMPANY'] = portfolio_lookup[isin]['MANAGEMENT_COMPANY']
    else:
        event_info['ASSET'] = 'N/A'
        event_info['PORTFOLIO'] = 'N/A'
        event_info['MANAGEMENT_COMPANY'] = ''
    
    events_by_day[day].append(event_info)

# --- ОТОБРАЖЕНИЕ КАЛЕНДАРЯ ---
cal = calendar.monthcalendar(year, month)

# Шапка
cols = st.columns(7)
for i, day_name in enumerate(['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']):
    cols[i].markdown(f"**{day_name}**")

# Ячейки календаря
for week in cal:
    cols = st.columns(7)
    for i, day in enumerate(week):
        with cols[i]:
            if day == 0:
                st.empty()
            else:
                if day in events_by_day:
                    # Получаем цвет для первой УК
                    mgmt = ''
                    for e in events_by_day[day]:
                        if e.get('MANAGEMENT_COMPANY'):
                            mgmt = e['MANAGEMENT_COMPANY']
                            break
                    color = uk_colors.get(mgmt, default_color) if mgmt else default_color
                    
                    # Кнопка дня
                    if st.button(
                        str(day),
                        key=f"day_{year}_{month}_{day}",
                        use_container_width=True,
                        type="secondary"
                    ):
                        st.session_state.selected_day = day
                        st.session_state.selected_events = events_by_day[day]
                    
                    # Цветная полоска
                    st.markdown(
                        f'<div style="background-color: {color}; height: 3px; border-radius: 2px;"></div>',
                        unsafe_allow_html=True
                    )
                    st.caption(f"{len(events_by_day[day])}")
                else:
                    st.write(f"**{day}**")

# --- ДЕТАЛИ ПО ВЫБРАННОМУ ДНЮ ---
if hasattr(st.session_state, 'selected_day') and st.session_state.selected_day:
    st.divider()
    st.markdown(f"### 📋 {st.session_state.selected_day} {calendar.month_name[month]} {year}")
    
    if st.session_state.selected_events:
        # Только ASSET и PORTFOLIO
        data = []
        for event in st.session_state.selected_events:
            data.append({
                'ASSET': event.get('ASSET', 'N/A'),
                'PORTFOLIO': event.get('PORTFOLIO', 'N/A')
            })
        
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True, hide_index=True)
