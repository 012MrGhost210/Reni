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
    page_title="📅 Календарь купонов",
    page_icon="📅",
    layout="wide"
)

# --- Загрузка данных ---
@st.cache_data
def load_data(file_path):
    """Загружает данные из файла Coupon_events.xlsx"""
    try:
        if not os.path.exists(file_path):
            st.error(f"Файл не найден: {file_path}")
            return pd.DataFrame()
        
        df = pd.read_excel(file_path)
        
        # Преобразуем дату
        df['DATE'] = pd.to_datetime(df['DATE'], format='%d.%m.%Y', errors='coerce')
        df = df.dropna(subset=['DATE'])
        
        # Сортируем по дате
        df = df.sort_values('DATE')
        
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
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
    st.error("Нет данных. Проверьте путь к файлу.")
    st.stop()

st.title("📅 Календарь купонов")

# --- Боковая панель с фильтрами ---
with st.sidebar:
    st.header("🎯 Фильтры")
    
    # Получаем список всех портфелей
    all_portfolios = sorted(df['PORTFOLIO'].unique().tolist())
    
    # Выбор портфелей
    selected_portfolios = st.multiselect(
        "Выберите портфели",
        options=all_portfolios,
        default=all_portfolios,
        help="Выберите один или несколько портфелей"
    )
    
    st.divider()
    
    # Выбор УК
    all_uk = sorted(df['MANAGEMENT_COMPANY'].unique().tolist())
    selected_uk = st.multiselect(
        "Выберите управляющие компании",
        options=all_uk,
        default=all_uk,
        help="Выберите одну или несколько УК"
    )
    
    st.divider()
    
    # Выбор месяца - ограничиваем только текущим месяцем и прошлыми
    today = datetime.now()
    current_year = today.year
    current_month = today.month
    
    # Доступные года (от 2020 до текущего)
    available_years = list(range(2020, current_year + 1))
    year = st.selectbox(
        "Год", 
        available_years, 
        index=len(available_years) - 1
    )
    
    # Доступные месяцы
    if year == current_year:
        available_months = list(range(1, current_month + 1))
    else:
        available_months = list(range(1, 13))
    
    month = st.selectbox(
        "Месяц", 
        available_months,
        format_func=lambda x: calendar.month_name[x],
        index=len(available_months) - 1
    )
    
    st.divider()
    
    # Статистика
    st.markdown("### 📊 Статистика")
    filtered_df = df[
        (df['PORTFOLIO'].isin(selected_portfolios)) &
        (df['MANAGEMENT_COMPANY'].isin(selected_uk))
    ]
    st.metric("Всего купонов", len(filtered_df))
    st.metric("Уникальных ISIN", filtered_df['ISIN'].nunique())
    st.metric("Уникальных портфелей", filtered_df['PORTFOLIO'].nunique())

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
        'NAME': event.get('NAME', '')
    })

# --- ОТОБРАЖЕНИЕ КАЛЕНДАРЯ ---
cal = calendar.monthcalendar(year, month)

# Шапка
cols = st.columns(7)
for i, day_name in enumerate(['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']):
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
                    
                    # Кнопка дня - только номер
                    if st.button(
                        str(day),
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
        st.markdown(f"### 📋 Купоны на {selected} {calendar.month_name[month]} {year}")
        
        if events_by_day[selected]:
            # Подготовка данных для таблицы
            data = []
            for event in events_by_day[selected]:
                data.append({
                    'Актив': event['ASSET'],
                    'Портфель': event['PORTFOLIO'],
                    'Управляющая компания': event['MANAGEMENT_COMPANY'],
                    'ISIN': event['ISIN'],
                    'Название': event['NAME']
                })
            
            df_day = pd.DataFrame(data)
            
            # Показываем таблицу
            st.dataframe(
                df_day,
                use_container_width=True,
                hide_index=True,
                column_config={
                    'ISIN': st.column_config.TextColumn(
                        width='small'
                    ),
                    'Название': st.column_config.TextColumn(
                        width='medium'
                    )
                }
            )
            
            # Статистика по портфелям
            st.caption(f"Всего купонов в этот день: {len(events_by_day[selected])}")
            
            # Разбивка по портфелям
            portfolio_counts = {}
            for event in events_by_day[selected]:
                p = event['PORTFOLIO']
                portfolio_counts[p] = portfolio_counts.get(p, 0) + 1
            
            if portfolio_counts:
                st.caption("Разбивка по портфелям:")
                for p, count in portfolio_counts.items():
                    color = get_portfolio_color(p, list(portfolio_counts.keys()).index(p))
                    st.markdown(
                        f'<div style="display: flex; align-items: center; gap: 8px; margin: 2px 0;">'
                        f'<div style="background-color: {color}; width: 12px; height: 12px; border-radius: 3px;"></div>'
                        f'<span>{p}: {count} купон(ов)</span>'
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
    st.markdown("### 🎨 Легенда")
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
st.caption(f"📊 Источник данных: {os.path.basename(DATA_FILE_PATH)} | Всего записей: {len(df)}")
