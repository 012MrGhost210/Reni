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

def get_portfolio_color(portfolio_name, idx):
    """Генерирует цвет для портфеля на основе индекса"""
    colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#FF8A5C', '#A8D8EA', '#FFD93D', '#6BCB77',
        '#FF9FF3', '#54A0FF', '#5F27CD', '#FF9F43', '#00D2D3'
    ]
    return colors[idx % len(colors)]

# --- Загружаем файлы ---
calendar_df, cal_err = load_calendar(CALENDAR_FILE_PATH)
portfolio_df, port_err = load_portfolio(PORTFOLIO_FILE_PATH)

if cal_err:
    st.error(f"Calendar error: {cal_err}")
    st.stop()

if port_err:
    st.warning(f"Portfolio error: {port_err}")

st.title("📅 Coupon Payment Calendar")

# --- Создаем lookup для портфеля по ISIN ---
portfolio_lookup = {}
all_portfolios = set()
if not portfolio_df.empty:
    for _, row in portfolio_df.iterrows():
        isin = row['ISIN'].strip()
        portfolio_lookup[isin] = {
            'ASSET': row.get('ASSET', ''),
            'PORTFOLIO': row.get('PORTFOLIO', ''),
            'MANAGEMENT_COMPANY': row.get('MANAGEMENT_COMPANY', '')
        }
        if row.get('PORTFOLIO'):
            all_portfolios.add(row.get('PORTFOLIO'))

# --- Боковая панель с фильтрами ---
with st.sidebar:
    st.header("🎯 Filters")
    
    # Выбор портфелей
    portfolio_options = sorted(list(all_portfolios))
    if portfolio_options:
        selected_portfolios = st.multiselect(
            "Select Portfolios",
            options=portfolio_options,
            default=portfolio_options,
            help="Choose one or multiple portfolios"
        )
    else:
        selected_portfolios = []
        st.warning("No portfolios found")
    
    st.divider()
    
    # Выбор месяца
    year = st.selectbox("Year", [2026, 2027, 2028], index=0)
    month = st.selectbox("Month", range(1, 13), 
                        format_func=lambda x: calendar.month_name[x],
                        index=datetime.now().month - 1)

# --- Основной контент ---
st.markdown(f"## {calendar.month_name[month]} {year}")

# Фильтруем события за месяц
first_day = date(year, month, 1)
last_day = date(year, month, calendar.monthrange(year, month)[1])
month_events = calendar_df[
    (calendar_df['DATE'].dt.date >= first_day) & 
    (calendar_df['DATE'].dt.date <= last_day)
]

# Группируем события по дням с учетом выбранных портфелей
events_by_day = {}
all_events_by_day = {}  # Все события для отображения в календаре

for _, event in month_events.iterrows():
    day = event['DATE'].day
    isin = event['ISIN'].strip()
    
    # Получаем данные из портфеля
    portfolio_info = portfolio_lookup.get(isin, {})
    portfolio = portfolio_info.get('PORTFOLIO', '')
    
    # Если портфель не выбран - пропускаем
    if selected_portfolios and portfolio not in selected_portfolios:
        continue
    
    if day not in events_by_day:
        events_by_day[day] = []
    
    events_by_day[day].append({
        'ISIN': isin,
        'NAME': event.get('NAME', ''),
        'ASSET': portfolio_info.get('ASSET', ''),
        'PORTFOLIO': portfolio,
        'PAYMENT': event.get('PAYMENT_RUB', 0),
        'MANAGEMENT_COMPANY': portfolio_info.get('MANAGEMENT_COMPANY', '')
    })

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
                    # Собираем портфели для этого дня
                    portfolios = set()
                    for e in events_by_day[day]:
                        if e.get('PORTFOLIO'):
                            portfolios.add(e['PORTFOLIO'])
                    
                    # Кнопка дня с количеством
                    if st.button(
                        f"{day}\n{len(events_by_day[day])}",
                        key=f"day_{year}_{month}_{day}",
                        use_container_width=True,
                        type="secondary"
                    ):
                        st.session_state.selected_day = day
                        st.session_state.selected_events = events_by_day[day]
                    
                    # Цветные полоски портфелей
                    if portfolios:
                        bars = ""
                        for idx, p in enumerate(portfolios):
                            color = get_portfolio_color(p, idx)
                            bars += f'<div style="background-color: {color}; height: 3px; border-radius: 2px; margin: 1px 0;"></div>'
                        st.markdown(bars, unsafe_allow_html=True)
                    else:
                        st.markdown(
                            '<div style="background-color: #D3D3D3; height: 3px; border-radius: 2px; margin: 1px 0;"></div>',
                            unsafe_allow_html=True
                        )
                else:
                    st.write(f"**{day}**")

# --- ДЕТАЛИ ПО ВЫБРАННОМУ ДНЮ ---
if hasattr(st.session_state, 'selected_day') and st.session_state.selected_day:
    selected = st.session_state.selected_day
    if selected in events_by_day:
        st.divider()
        st.markdown(f"### 📋 Coupons for {selected} {calendar.month_name[month]} {year}")
        
        if events_by_day[selected]:
            # Подготовка данных для таблицы - ТОЛЬКО активы с купонами
            data = []
            for event in events_by_day[selected]:
                data.append({
                    'Asset': event.get('ASSET', ''),
                    'Portfolio': event.get('PORTFOLIO', ''),
                    'Payment (RUB)': event.get('PAYMENT', 0),
                    'ISIN': event.get('ISIN', '')
                })
            
            df = pd.DataFrame(data)
            
            # Показываем таблицу
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    'Payment (RUB)': st.column_config.NumberColumn(
                        format="%.2f ₽",
                        help="Coupon payment amount"
                    ),
                    'ISIN': st.column_config.TextColumn(
                        width='small',
                        help="ISIN code"
                    )
                }
            )
            
            # Статистика по портфелям
            st.caption(f"Total coupons: {len(events_by_day[selected])}")
            
            # Breakdown по портфелям
            portfolio_counts = {}
            for event in events_by_day[selected]:
                p = event.get('PORTFOLIO', '')
                if p:
                    portfolio_counts[p] = portfolio_counts.get(p, 0) + 1
            
            if portfolio_counts:
                st.caption("Breakdown by Portfolio:")
                for p, count in portfolio_counts.items():
                    st.caption(f"• {p}: {count} coupon(s)")
    else:
        st.info("No coupons on this day")

# --- Легенда портфелей ---
# Собираем все портфели, которые есть в отфильтрованных данных
shown_portfolios = set()
for day_events in events_by_day.values():
    for e in day_events:
        if e.get('PORTFOLIO'):
            shown_portfolios.add(e['PORTFOLIO'])

if shown_portfolios:
    st.divider()
    st.markdown("### 🎨 Legend")
    cols = st.columns(min(len(shown_portfolios), 5))
    for idx, p in enumerate(sorted(shown_portfolios)):
        with cols[idx % len(cols)]:
            color = get_portfolio_color(p, idx)
            st.markdown(
                f'<div style="display: flex; align-items: center; gap: 8px;">'
                f'<div style="background-color: {color}; width: 16px; height: 16px; border-radius: 4px;"></div>'
                f'<span style="font-size: 13px;">{p}</span>'
                f'</div>',
                unsafe_allow_html=True
            )
