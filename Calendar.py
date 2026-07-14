NAV_DATE
PORTFOLIO
MANAGEMENT_COMPANY
ASSET
ISIN
ДУ «ТКБ Инвестмент Партнерс» 18-226-РЖ
Собственные средства СК РЖ (ГПБ_Г000-Б-28180)
ДУ «ТКБ Инвестмент Партнерс» 18-230-РЖ
ДУ «ТКБ Инвестмент Партнерс» 21-229-РЖ
ДУ «Спутник-УК» 190221/1 SPURZ 10
ДУ «Спутник-УК» 271210/2 SPURZ
ДУ «Спутник-УК» 020611/2 SPURZ 2
ДУ «Спутник-УК» 260716/1 SPURZ 5
ДУ «Спутник-УК» 020611/3 SPURZ 3
ДУ «Спутник-УК» 050925/1 SPURZ 15
Собственные средства СК РЖ (ГПБ_Г000-Б-50817)
Собственные средства СК РЖ (ГПБ_Г000-Б-514660)
ДУ "УК Первая"_УК-11/2026
ДУ «Спутник-УК» 020611/1 SPURZ 1
ДУ «ТКБ Инвестмент Партнерс» 21-228-РЖ
Собственные средства СК РЖ REZHS
ДУ «Спутник-УК» 220223/2 SPURZ 14
ДУ «Райффайзен УК» 256-03.1ДУ/24
ДУ «Райффайзен УК» 257-03.1ДУ/24
ДУ «ТКБ Инвестмент Партнерс» 18-225-РЖ
ДУ «Спутник-УК» 081121/1 SPURZ 11
ДУ «Спутник-УК» 081121/2 SPURZ 12

ТКБ ИНВЕСТМЕНТ ПАРТНЕРС (АО)
ООО СК РЕНЕССАНС ЖИЗНЬ 
УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ АО
ПЕРВАЯ АО УК
УК РАЙФФАЙЗЕН ООО

import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, date
import os

# ============================================================
# НАСТРАИВАЕМЫЕ ПАРАМЕТРЫ - отредактируйте пути к вашим файлам
# ============================================================
# Путь к файлу календаря (XLSX)
CALENDAR_FILE_PATH = "Календарь (3).xlsx"

# Путь к файлу портфеля X (XLSX) - должен содержать колонки: NAV_DATE, PORTFOLIO, MANAGEMENT_COMPANY, ASSET, ISIN
PORTFOLIO_FILE_PATH = "portfolio_x.xlsx"
# ============================================================

# --- Page config ---
st.set_page_config(
    page_title="📅 Coupon Payment Calendar",
    page_icon="📅",
    layout="wide"
)

# --- Helper Functions ---
def load_calendar_from_path(file_path):
    """Load calendar data from a file path (XLSX)"""
    try:
        if not os.path.exists(file_path):
            return pd.DataFrame(), f"File not found: {file_path}"
        
        # Читаем XLSX файл
        df = pd.read_excel(file_path, skiprows=3)
        
        # Переименовываем колонки
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
        return pd.DataFrame(), f"Error loading calendar file: {e}"

def load_portfolio_from_path(file_path):
    """Load portfolio data from a file path (XLSX)"""
    try:
        if not os.path.exists(file_path):
            return pd.DataFrame(), f"File not found: {file_path}"
        
        # Читаем XLSX файл
        df = pd.read_excel(file_path)
        
        # Проверяем только те колонки, которые указал пользователь
        required_cols = ['NAV_DATE', 'PORTFOLIO', 'MANAGEMENT_COMPANY', 'ASSET', 'ISIN']
        if all(col in df.columns for col in required_cols):
            df['NAV_DATE'] = pd.to_datetime(df['NAV_DATE'], errors='coerce')
            df['ISIN'] = df['ISIN'].astype(str).str.strip()
            return df, None
        else:
            missing = [col for col in required_cols if col not in df.columns]
            return pd.DataFrame(), f"Missing columns: {missing}"
    except Exception as e:
        return pd.DataFrame(), f"Error loading portfolio file: {e}"

@st.cache_data
def load_calendar_file(uploaded_file):
    """Load calendar data from uploaded file (XLSX)"""
    if uploaded_file is not None:
        try:
            # Читаем XLSX файл
            df = pd.read_excel(uploaded_file, skiprows=3)
            
            # Переименовываем колонки
            df.columns = ['ISIN', 'NAME', 'VOLUME', 'DATE', 'NOMINAL', 'CURRENCY', 
                         'OUTSTANDING_NOMINAL', 'COUPON_RATE', 'PAYMENT', 'PAYMENT_RUB']
            
            df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
            df = df.dropna(subset=['DATE'])
            df['ISIN'] = df['ISIN'].astype(str).str.strip()
            df = df[df['ISIN'] != 'null']
            df = df[df['ISIN'] != '']
            df['PAYMENT_RUB'] = pd.to_numeric(df['PAYMENT_RUB'], errors='coerce')
            
            return df
        except Exception as e:
            st.error(f"Error loading calendar file: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

@st.cache_data
def load_portfolio_file(uploaded_file):
    """Load portfolio data from uploaded file (XLSX)"""
    if uploaded_file is not None:
        try:
            # Читаем XLSX файл
            df = pd.read_excel(uploaded_file)
            
            # Проверяем только те колонки, которые указал пользователь
            required_cols = ['NAV_DATE', 'PORTFOLIO', 'MANAGEMENT_COMPANY', 'ASSET', 'ISIN']
            if all(col in df.columns for col in required_cols):
                df['NAV_DATE'] = pd.to_datetime(df['NAV_DATE'], errors='coerce')
                df['ISIN'] = df['ISIN'].astype(str).str.strip()
                return df
            else:
                missing = [col for col in required_cols if col not in df.columns]
                st.error(f"Missing columns: {missing}")
                st.write("Found columns:", df.columns.tolist())
                return pd.DataFrame()
        except Exception as e:
            st.error(f"Error loading portfolio file: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def create_calendar_grid(year, month, calendar_df, portfolio_df):
    """Create calendar grid data with coupon events"""
    # Portfolio lookup dictionary using ISIN
    portfolio_lookup = {}
    if not portfolio_df.empty:
        for _, row in portfolio_df.iterrows():
            isin = row['ISIN']
            portfolio_lookup[isin] = {
                'MANAGEMENT_COMPANY': row.get('MANAGEMENT_COMPANY', 'N/A'),
                'ASSET': row.get('ASSET', 'N/A'),
                'PORTFOLIO': row.get('PORTFOLIO', 'N/A')
            }
    
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])
    
    month_events = calendar_df[
        (calendar_df['DATE'].dt.date >= first_day) & 
        (calendar_df['DATE'].dt.date <= last_day)
    ]
    
    cal = calendar.monthcalendar(year, month)
    
    calendar_matrix = []
    for week in cal:
        week_data = []
        for day in week:
            if day == 0:
                week_data.append({
                    'day': 0,
                    'events': [],
                    'has_events': False,
                    'management_companies': []
                })
            else:
                current_date = date(year, month, day)
                day_events = month_events[month_events['DATE'].dt.date == current_date]
                
                events_list = []
                management_list = []
                
                for _, event in day_events.iterrows():
                    isin = event['ISIN']
                    event_info = {
                        'ISIN': isin,
                        'NAME': event.get('NAME', 'N/A'),
                        'PAYMENT': event.get('PAYMENT_RUB', 0),
                        'MANAGEMENT_COMPANY': 'N/A',
                        'ASSET': 'N/A',
                        'PORTFOLIO': 'N/A'
                    }
                    
                    if isin in portfolio_lookup:
                        event_info['MANAGEMENT_COMPANY'] = portfolio_lookup[isin]['MANAGEMENT_COMPANY']
                        event_info['ASSET'] = portfolio_lookup[isin]['ASSET']
                        event_info['PORTFOLIO'] = portfolio_lookup[isin]['PORTFOLIO']
                        management_list.append(portfolio_lookup[isin]['MANAGEMENT_COMPANY'])
                    
                    events_list.append(event_info)
                
                week_data.append({
                    'day': day,
                    'events': events_list,
                    'has_events': len(events_list) > 0,
                    'management_companies': list(set(management_list))
                })
        calendar_matrix.append(week_data)
    
    return calendar_matrix

def get_company_color_map(calendar_matrix):
    """Extract unique management companies and assign colors"""
    companies = set()
    for week in calendar_matrix:
        for day in week:
            if day['day'] != 0:
                companies.update(day['management_companies'])
    
    companies.discard('N/A')
    
    color_map = {
        'ТКБ ИНВЕСТМЕНТ ПАРТНЕРС (АО)': '#FF6B6B',
        'ООО СК РЕНЕССАНС ЖИЗНЬ': '#4ECDC4',
        'УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ АО': '#45B7D1',
        'ПЕРВАЯ АО УК': '#96CEB4',
        'УК РАЙФФАЙЗЕН ООО': '#FFD93D',
    }
    
    default_colors = ['#FF8A5C', '#A8D8EA', '#DDA0DD', '#FFEAA7', '#6BCB77']
    default_idx = 0
    
    for company in sorted(companies):
        if company not in color_map:
            color_map[company] = default_colors[default_idx % len(default_colors)]
            default_idx += 1
    
    return color_map

# --- Main App ---
st.title("📅 Coupon Payment Calendar")

# Sidebar
with st.sidebar:
    st.header("📂 File Configuration")
    
    st.markdown("### Current File Paths")
    st.code(f"Calendar: {CALENDAR_FILE_PATH}", language="text")
    st.code(f"Portfolio: {PORTFOLIO_FILE_PATH}", language="text")
    st.caption("💡 To change paths, edit the variables at the top of the script")
    
    st.divider()
    
    st.markdown("### 📤 Upload Files (Optional)")
    st.caption("Upload files manually if you don't want to use the configured paths")
    
    calendar_file = st.file_uploader(
        "Upload calendar XLSX",
        type=['xlsx'],
        key="calendar_uploader"
    )
    
    portfolio_file = st.file_uploader(
        "Upload portfolio XLSX",
        type=['xlsx'],
        key="portfolio_uploader"
    )
    
    st.divider()
    
    st.markdown("### 📊 Controls")
    
    current_year = datetime.now().year
    current_month = datetime.now().month
    
    year = st.selectbox(
        "Select Year",
        options=[2026, 2027, 2028],
        index=0
    )
    
    month = st.selectbox(
        "Select Month",
        options=[(i, calendar.month_name[i]) for i in range(1, 13)],
        format_func=lambda x: x[1],
        index=current_month - 1
    )[0]
    
    st.divider()
    
    st.markdown("### 🏢 Management Companies")
    st.caption("Expected in File X:")
    mgmt_companies = [
        "ТКБ ИНВЕСТМЕНТ ПАРТНЕРС (АО)",
        "ООО СК РЕНЕССАНС ЖИЗНЬ",
        "УК СПУТНИК - УПРАВЛЕНИЕ КАПИТАЛОМ АО",
        "ПЕРВАЯ АО УК",
        "УК РАЙФФАЙЗЕН ООО"
    ]
    for mgmt in mgmt_companies:
        st.caption(f"• {mgmt}")
    
    st.divider()
    
    st.markdown("### ℹ️ Instructions")
    st.info("""
    1. Edit file paths at the top of the script
    2. Or upload files manually below
    3. Select year and month
    4. Click on any highlighted day to see details
    """)

# Main content
calendar_df = pd.DataFrame()
portfolio_df = pd.DataFrame()
use_uploaded = False

# Try to load from configured paths
if os.path.exists(CALENDAR_FILE_PATH):
    calendar_df, cal_error = load_calendar_from_path(CALENDAR_FILE_PATH)
    if cal_error:
        st.sidebar.warning(f"Calendar error: {cal_error}")

if os.path.exists(PORTFOLIO_FILE_PATH):
    portfolio_df, port_error = load_portfolio_from_path(PORTFOLIO_FILE_PATH)
    if port_error:
        st.sidebar.warning(f"Portfolio error: {port_error}")

# If configured files didn't load, try uploaded files
if calendar_df.empty and calendar_file is not None:
    calendar_df = load_calendar_file(calendar_file)
    use_uploaded = True

if portfolio_df.empty and portfolio_file is not None:
    portfolio_df = load_portfolio_file(portfolio_file)
    use_uploaded = True

# Show data source
if use_uploaded:
    st.caption("📤 Using uploaded files")
elif not calendar_df.empty and not portfolio_df.empty:
    st.caption(f"📁 Using files from: {os.path.dirname(CALENDAR_FILE_PATH)}")

if not calendar_df.empty:
    calendar_matrix = create_calendar_grid(year, month, calendar_df, portfolio_df)
    company_color_map = get_company_color_map(calendar_matrix)
    
    st.markdown(f"## {calendar.month_name[month]} {year}")
    
    days_of_week = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    
    cols = st.columns(7)
    for i, day_name in enumerate(days_of_week):
        cols[i].markdown(f"**{day_name}**")
    
    for week in calendar_matrix:
        cols = st.columns(7)
        for i, day_data in enumerate(week):
            with cols[i]:
                if day_data['day'] == 0:
                    st.empty()
                else:
                    if day_data['has_events']:
                        if day_data['management_companies']:
                            main_company = day_data['management_companies'][0]
                            color = company_color_map.get(main_company, '#D3D3D3')
                        else:
                            color = '#D3D3D3'
                        
                        if st.button(
                            f"**{day_data['day']}**\n{len(day_data['events'])} 📌",
                            key=f"day_{year}_{month}_{day_data['day']}",
                            use_container_width=True,
                            type="secondary"
                        ):
                            st.session_state.selected_day = {
                                'year': year,
                                'month': month,
                                'day': day_data['day'],
                                'events': day_data['events']
                            }
                        
                        st.markdown(
                            f'<div style="background-color: {color}; height: 4px; border-radius: 2px;"></div>',
                            unsafe_allow_html=True
                        )
                        
                        if day_data['management_companies']:
                            companies_text = ', '.join(day_data['management_companies'][:2])
                            if len(day_data['management_companies']) > 2:
                                companies_text += f' +{len(day_data["management_companies"]) - 2}'
                            st.caption(companies_text)
                    else:
                        st.write(f"**{day_data['day']}**")
    
    # Display selected day details
    if hasattr(st.session_state, 'selected_day') and st.session_state.selected_day:
        selected = st.session_state.selected_day
        if selected['year'] == year and selected['month'] == month:
            st.divider()
            st.markdown(f"### 📋 Details for {selected['day']} {calendar.month_name[month]} {year}")
            
            if selected['events']:
                events_data = []
                for event in selected['events']:
                    events_data.append({
                        'ASSET': event.get('ASSET', 'N/A'),
                        'PORTFOLIO': event.get('PORTFOLIO', 'N/A'),
                        'MANAGEMENT_COMPANY': event.get('MANAGEMENT_COMPANY', 'N/A'),
                        'ISIN': event.get('ISIN', 'N/A'),
                        'Payment (RUB)': event.get('PAYMENT', 0),
                        'Name': event.get('NAME', 'N/A')[:50] + '...' if len(event.get('NAME', '')) > 50 else event.get('NAME', 'N/A')
                    })
                
                df_events = pd.DataFrame(events_data)
                
                st.dataframe(
                    df_events,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        'Payment (RUB)': st.column_config.NumberColumn(format="%.2f ₽"),
                        'ISIN': st.column_config.TextColumn(width='small'),
                    }
                )
                
                st.caption(f"Total coupons on this day: {len(selected['events'])}")
                
                mgmt_counts = df_events['MANAGEMENT_COMPANY'].value_counts()
                if not mgmt_counts.empty:
                    st.caption("Breakdown by Management Company:")
                    for mgmt, count in mgmt_counts.items():
                        if mgmt != 'N/A':
                            st.caption(f"• {mgmt}: {count} coupon(s)")
            else:
                st.info("No coupon events for this day")
    
    # Summary statistics
    st.divider()
    col1, col2, col3 = st.columns(3)
    
    total_events = sum(1 for week in calendar_matrix for day in week if day['day'] != 0 for _ in day['events'])
    total_days_with_events = sum(1 for week in calendar_matrix for day in week if day['day'] != 0 and day['has_events'])
    
    mgmt_companies = set()
    for week in calendar_matrix:
        for day in week:
            if day['day'] != 0:
                mgmt_companies.update(day['management_companies'])
    mgmt_companies.discard('N/A')
    
    col1.metric("📌 Total Coupons", total_events)
    col2.metric("📅 Days with Coupons", total_days_with_events)
    col3.metric("🏢 Management Companies", len(mgmt_companies))
    
    if company_color_map:
        st.markdown("### 🎨 Legend")
        legend_cols = st.columns(min(len(company_color_map), 4))
        for i, (company, color) in enumerate(company_color_map.items()):
            with legend_cols[i % len(legend_cols)]:
                st.markdown(
                    f'<div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">'
                    f'<div style="background-color: {color}; width: 16px; height: 16px; border-radius: 4px; flex-shrink: 0;"></div>'
                    f'<span style="font-size: 12px;">{company}</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )
    
else:
    st.warning("No data loaded. Please check file paths or upload files manually.")

st.divider()
st.caption("📊 Coupon Calendar Dashboard | Edit file paths at the top of the script")
