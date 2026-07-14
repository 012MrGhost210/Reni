import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, date
import os

# ============================================================
# НАСТРАИВАЕМЫЕ ПАРАМЕТРЫ
# ============================================================
DATA_FILE_PATH = r"M:\Финансовый департамент\Treasury\3. ЗАКРЫТИЕ\Отчеты УК\Сводные отчеты УК\Coupon_events.xlsx"
# ============================================================

st.set_page_config(
    page_title="📅 Coupon Payment Calendar",
    page_icon="📅",
    layout="wide"
)

# --- Загрузка данных ---
@st.cache_data
def load_data(file_path):
    """Загружает данные из файла Coupon_events.xlsx"""
    try:
        if not os.path.exists(file_path):
            st.error(f"File not found: {file_path}")
            return pd.DataFrame()
        
        df = pd.read_excel(file_path)
        
        # Преобразуем дату
        df['DATE'] = pd.to_datetime(df['DATE'], format='%d.%m.%Y', errors='coerce')
        df = df.dropna(subset=['DATE'])
        
        # Сортируем по дате
        df = df.sort_values('DATE')
        
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()

# --- Функция для цветов портфелей ---
def get_portfolio_color(portfolio_name, idx):
    colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#FF8A5C', '#A8D8EA', '#FFD93D', '#6BCB77',
        '#FF9FF3', '#54A0FF', '#5F27CD', '#FF9F43', '#00D2D3',
        '#F368E0', '#FFC048', '#74B9FF', '#55EFC4', '#FD79A8'
    ]
    return colors[idx % len(colors)]

# --- Загружаем данные ---
df = load_data(DATA_FILE_PATH)

if df.empty:
    st.error("No data loaded. Please check the file path.")
    st.stop()

st.title("📅 Coupon Payment Calendar")

# --- Боковая панель с фильтрами ---
with st.sidebar:
    st.header("🎯 Filters")
    
    # Получаем список всех портфелей
    all_portfolios = sorted(df['PORTFOLIO'].unique().tolist())
    
    # Выбор портфелей
    selected_portfolios = st.multiselect(
        "Select Portfolios",
        options=all_portfolios,
        default=all_portfolios,
        help="Choose one or multiple portfolios"
    )
    
    st.divider()
    
    # Выбор УК
    all_uk = sorted(df['MANAGEMENT_COMPANY'].unique().tolist())
    selected_uk = st.multiselect(
        "Select Management Companies",
        options=all_uk,
        default=all_uk,
        help="Choose one or multiple management companies"
    )
    
    st.divider()
    
    # Выбор месяца
    current_year = datetime.now().year
    year = st.selectbox("Year", [2026, 2027, 2028], index=0)
    month = st.selectbox("Month", range(1, 13), 
                        format_func=lambda x: calendar.month_name[x],
                        index=datetime.now().month - 1)
    
    st.divider()
    
    # Статистика
    st.markdown("### 📊 Statistics")
    filtered_df = df[
        (df['PORTFOLIO'].isin(selected_portfolios)) &
        (df['MANAGEMENT_COMPANY'].isin(selected_uk))
    ]
    st.metric("Total Coupons", len(filtered_df))
    st.metric("Unique ISINs", filtered_df['ISIN'].nunique())
    st.metric("Unique Portfolios", filtered_df['PORTFOLIO'].nunique())

# --- Основной контент ---
st.markdown(f"## {calendar.month_name[month]} {year}")

# Фильтруем данные
filtered_df = df[
    (df['PORTFOLIO'].isin(selected_portfolios)) &
    (df['MANAGEMENT_COMPANY'].isin(selected_uk))
]

# Фильтруем по месяцу
month_start = date(year, month, 1)
month_end = date(year, month, calendar.monthrange(year, month)[1])
month_df = filtered_df[
    (filtered_df['DATE'].dt.date >= month_start) & 
    (filtered_df['DATE'].dt.date <= month_end)
]

# Группируем по дням
events_by_day = {}
for _, event in month_df.iterrows():
    day = event['DATE'].day
    if day not in events_by_day:
        events_by_day[day] = []
    
    events_by_day[day].append({
        'ISIN': event['ISIN'],
        'ASSET': event['ASSET'],
        'PORTFOLIO': event['PORTFOLIO'],
        'MANAGEMENT_COMPANY': event['MANAGEMENT_COMPANY'],
        'PAYMENT_RUB': event['PAYMENT_RUB'],
        'NAME': event.get('NAME', '')
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
                        portfolios.add(e['PORTFOLIO'])
                    
                    # Кнопка дня
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
                    st.write(f"**{day}**")

# --- ДЕТАЛИ ПО ВЫБРАННОМУ ДНЮ ---
if hasattr(st.session_state, 'selected_day') and st.session_state.selected_day:
    selected = st.session_state.selected_day
    if selected in events_by_day:
        st.divider()
        st.markdown(f"### 📋 Coupons for {selected} {calendar.month_name[month]} {year}")
        
        if events_by_day[selected]:
            # Подготовка данных для таблицы
            data = []
            for event in events_by_day[selected]:
                data.append({
                    'ASSET': event['ASSET'],
                    'PORTFOLIO': event['PORTFOLIO'],
                    'MANAGEMENT_COMPANY': event['MANAGEMENT_COMPANY'],
                    'Payment (RUB)': event['PAYMENT_RUB'],
                    'ISIN': event['ISIN']
                })
            
            df_day = pd.DataFrame(data)
            
            # Показываем таблицу
            st.dataframe(
                df_day,
                use_container_width=True,
                hide_index=True,
                column_config={
                    'Payment (RUB)': st.column_config.NumberColumn(
                        format="%.2f ₽"
                    ),
                    'ISIN': st.column_config.TextColumn(
                        width='small'
                    )
                }
            )
            
            # Статистика по портфелям
            st.caption(f"Total coupons: {len(events_by_day[selected])}")
            
            # Breakdown по портфелям
            portfolio_counts = {}
            for event in events_by_day[selected]:
                p = event['PORTFOLIO']
                portfolio_counts[p] = portfolio_counts.get(p, 0) + 1
            
            if portfolio_counts:
                st.caption("Breakdown by Portfolio:")
                for p, count in portfolio_counts.items():
                    color = get_portfolio_color(p, list(portfolio_counts.keys()).index(p))
                    st.markdown(
                        f'<div style="display: flex; align-items: center; gap: 8px; margin: 2px 0;">'
                        f'<div style="background-color: {color}; width: 12px; height: 12px; border-radius: 3px;"></div>'
                        f'<span>{p}: {count} coupon(s)</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

# --- Легенда портфелей ---
# Собираем все портфели в отфильтрованных данных
shown_portfolios = set()
for day_events in events_by_day.values():
    for e in day_events:
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

# --- Информация о данных ---
st.divider()
st.caption(f"📊 Data source: {os.path.basename(DATA_FILE_PATH)} | Total records: {len(df)}")
