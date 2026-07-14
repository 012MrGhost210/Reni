import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, date
import os
import random

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
        df['PAYMENT_RUB'] = pd.to_numeric(df['PAYMENT_RUB'], errors='coerce')
        return df, None
    except Exception as e:
        return pd.DataFrame(), str(e)

def load_portfolio(file_path):
    try:
        if not os.path.exists(file_path):
            return pd.DataFrame(), f"File not found: {file_path}"
        df = pd.read_excel(file_path)
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

# --- Функция для генерации цветов портфелей ---
def get_portfolio_color(portfolio_name):
    colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#FF8A5C', '#A8D8EA', '#FFD93D', '#6BCB77',
        '#FF9FF3', '#54A0FF', '#5F27CD', '#FF9F43', '#00D2D3',
        '#F368E0', '#FFC048', '#74B9FF', '#55EFC4', '#FD79A8'
    ]
    # Используем хеш для стабильного цвета
    hash_val = hash(portfolio_name) % len(colors)
    return colors[hash_val]

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
        isin = row['ISIN'].strip()
        portfolio_lookup[isin] = {
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

# Группируем события по дням
events_by_day = {}
for _, event in month_events.iterrows():
    day = event['DATE'].day
    if day not in events_by_day:
        events_by_day[day] = []
    
    isin = event['ISIN'].strip()
    event_info = {
        'ISIN': isin,
        'NAME': event.get('NAME', ''),
        'PAYMENT': event.get('PAYMENT_RUB', 0),
    }
    
    # Ищем в портфеле
    if isin in portfolio_lookup:
        event_info['ASSET'] = portfolio_lookup[isin]['ASSET']
        event_info['PORTFOLIO'] = portfolio_lookup[isin]['PORTFOLIO']
        event_info['MANAGEMENT_COMPANY'] = portfolio_lookup[isin]['MANAGEMENT_COMPANY']
        event_info['FOUND'] = True
    else:
        # Если не нашли ISIN - помечаем как неизвестный
        event_info['ASSET'] = 'Not found'
        event_info['PORTFOLIO'] = 'Unknown'
        event_info['MANAGEMENT_COMPANY'] = 'Unknown'
        event_info['FOUND'] = False
    
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
                    # Собираем все портфели для этого дня
                    portfolios = set()
                    for e in events_by_day[day]:
                        if e.get('PORTFOLIO') and e['PORTFOLIO'] != 'Unknown':
                            portfolios.add(e['PORTFOLIO'])
                    
                    # Если есть портфели - показываем цветные полоски
                    if portfolios:
                        # Создаем контейнер для полосок
                        color_bars = []
                        for p in portfolios:
                            color = get_portfolio_color(p)
                            color_bars.append(
                                f'<div style="background-color: {color}; height: 4px; border-radius: 2px; margin: 1px 0;"></div>'
                            )
                        
                        # Кнопка дня
                        if st.button(
                            str(day),
                            key=f"day_{year}_{month}_{day}",
                            use_container_width=True,
                            type="secondary"
                        ):
                            st.session_state.selected_day = day
                            st.session_state.selected_events = events_by_day[day]
                        
                        # Показываем полоски портфелей
                        st.markdown(
                            f'<div style="margin-top: 2px;">{"".join(color_bars)}</div>',
                            unsafe_allow_html=True
                        )
                        
                        # Показываем количество купонов
                        st.caption(f"{len(events_by_day[day])} 📌")
                        
                        # Показываем названия портфелей (кратко)
                        if len(portfolios) <= 3:
                            st.caption(", ".join(list(portfolios)))
                        else:
                            st.caption(f"{len(portfolios)} portfolios")
                    else:
                        # Если портфелей нет - показываем серую полоску
                        if st.button(
                            str(day),
                            key=f"day_{year}_{month}_{day}",
                            use_container_width=True,
                            type="secondary"
                        ):
                            st.session_state.selected_day = day
                            st.session_state.selected_events = events_by_day[day]
                        
                        st.markdown(
                            f'<div style="background-color: #D3D3D3; height: 3px; border-radius: 2px; margin-top: 2px;"></div>',
                            unsafe_allow_html=True
                        )
                        st.caption(f"{len(events_by_day[day])} 📌")
                else:
                    st.write(f"**{day}**")

# --- ДЕТАЛИ ПО ВЫБРАННОМУ ДНЮ ---
if hasattr(st.session_state, 'selected_day') and st.session_state.selected_day:
    st.divider()
    st.markdown(f"### 📋 {st.session_state.selected_day} {calendar.month_name[month]} {year}")
    
    if st.session_state.selected_events:
        # Подготовка данных для таблицы
        data = []
        for event in st.session_state.selected_events:
            # Проверяем, что ASSET не пустой и не N/A
            asset = event.get('ASSET', '')
            if not asset or asset == 'Not found':
                asset = '⚠️ Not in portfolio'
            
            portfolio = event.get('PORTFOLIO', '')
            if not portfolio or portfolio == 'Unknown':
                portfolio = '⚠️ Not found'
            
            data.append({
                'Asset': asset,
                'Portfolio': portfolio,
                'Payment (RUB)': event.get('PAYMENT', 0) if event.get('PAYMENT', 0) > 0 else '-',
                'ISIN': event.get('ISIN', '')
            })
        
        df = pd.DataFrame(data)
        
        # Стилизация таблицы
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Payment (RUB)': st.column_config.NumberColumn(format="%.2f ₽"),
                'ISIN': st.column_config.TextColumn(width='small'),
            }
        )
        
        # Сводка по портфелям
        portfolios_count = {}
        for event in st.session_state.selected_events:
            p = event.get('PORTFOLIO', 'Unknown')
            if p and p != 'Unknown':
                portfolios_count[p] = portfolios_count.get(p, 0) + 1
        
        if portfolios_count:
            st.caption("📊 Breakdown by Portfolio:")
            for p, count in portfolios_count.items():
                color = get_portfolio_color(p)
                st.markdown(
                    f'<div style="display: flex; align-items: center; gap: 8px; margin: 2px 0;">'
                    f'<div style="background-color: {color}; width: 12px; height: 12px; border-radius: 3px;"></div>'
                    f'<span>{p}: {count} coupon(s)</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )
        
        # Предупреждение о ненайденных ISIN
        not_found = [e for e in st.session_state.selected_events if e.get('ASSET') == 'Not found']
        if not_found:
            st.warning(f"⚠️ {len(not_found)} ISIN(s) not found in portfolio")

# --- Легенда портфелей (все уникальные портфели за месяц) ---
st.divider()
st.markdown("### 🎨 Portfolio Legend")

# Собираем все портфели за месяц
all_portfolios = set()
for day_events in events_by_day.values():
    for e in day_events:
        p = e.get('PORTFOLIO')
        if p and p != 'Unknown':
            all_portfolios.add(p)

if all_portfolios:
    # Показываем в несколько колонок
    cols = st.columns(min(len(all_portfolios), 5))
    for i, p in enumerate(sorted(all_portfolios)):
        with cols[i % len(cols)]:
            color = get_portfolio_color(p)
            st.markdown(
                f'<div style="display: flex; align-items: center; gap: 8px; margin: 2px 0;">'
                f'<div style="background-color: {color}; width: 16px; height: 16px; border-radius: 4px;"></div>'
                f'<span style="font-size: 13px;">{p}</span>'
                f'</div>',
                unsafe_allow_html=True
            )
else:
    st.info("No portfolio data found for this month")
