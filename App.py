import streamlit as st
import subprocess
import sys
import os
import locale
import time
from datetime import datetime, timedelta

# Настройка страницы
st.set_page_config(
    page_title="🚀 Запуск скриптов",
    page_icon="🚀",
    layout="wide"
)

# ==================== КОНФИГУРАЦИЯ СКРИПТОВ ====================
# Основная группа скриптов
MAIN_SCRIPTS_CONFIG = {
    "SFT.py": "📝 СФТ Платеж",
    "rusfar.py": "📈 Обновление Rusfar",
    "INDEX.py": "📊 Обновление индексов",
    "simple.py": "📁 Скопировать файл Daily Income 2026",
    "OSVRG.py": "🎯ОСВ",
    "check.py": "⚙️ Проверить наличие актуальных файлов в диадоке",
    "SCHA.py": "✅СЧА для ИД РЖ"
}

# Дополнительная группа скриптов
EXTRA_SCRIPTS_CONFIG = {
    "zaprosstavok.py": "📨 Запрос ставок",
    "stop.py": "❌ Убрать заглушку"
}

# ==================== НАСТРОЙКИ ТАЙМАУТОВ ====================
# Можно настроить для каждого скрипта индивидуально (в секундах)
SCRIPT_TIMEOUTS = {
    # Основные скрипты
    "SFT.py": 0,
    "rusfar.py": 5,
    "INDEX.py": 5,
    "simple.py": 0,
    "OSVRG.py": 5,
    "check.py": 0,
    "SCHA.py": 40,
    
    # Дополнительные скрипты
    "zaprosstavok.py": 5,
    "stop.py": 300,
}

# Таймаут по умолчанию, если скрипт не указан в SCRIPT_TIMEOUTS
DEFAULT_TIMEOUT = 3
# ===============================================================

# Определяем системную кодировку
try:
    SYSTEM_ENCODING = locale.getpreferredencoding()
except:
    SYSTEM_ENCODING = 'cp1251'

st.title("🚀 Запуск Python скриптов")
st.markdown("---")

# Инициализация session state для хранения времени последнего нажатия
if 'last_click_time' not in st.session_state:
    st.session_state.last_click_time = {}

if 'button_cooldown' not in st.session_state:
    st.session_state.button_cooldown = {}

# Функция для проверки таймаута кнопки
def is_button_on_cooldown(button_key, cooldown_seconds):
    """
    Проверяет, находится ли кнопка в состоянии таймаута
    Возвращает (bool, float) - (на таймауте ли, сколько секунд осталось)
    """
    if button_key in st.session_state.last_click_time:
        last_click = st.session_state.last_click_time[button_key]
        time_diff = (datetime.now() - last_click).total_seconds()
        
        if time_diff < cooldown_seconds:
            remaining = round(cooldown_seconds - time_diff, 1)
            return True, remaining
    
    return False, 0

# Функция для обновления времени последнего нажатия
def update_last_click(button_key):
    st.session_state.last_click_time[button_key] = datetime.now()

# Функция для чтения файла инструкции
def read_instruction_file(script_name):
    """Читает содержимое txt файла с инструкцией"""
    # Заменяем расширение .py на .txt
    txt_file = script_name.replace('.py', '.txt')
    
    try:
        if os.path.exists(txt_file):
            with open(txt_file, 'r', encoding='utf-8') as f:
                return f.read()
        else:
            return f"❌ Файл инструкции {txt_file} не найден"
    except Exception as e:
        return f"❌ Ошибка при чтении инструкции: {str(e)}"

# Функция для создания кнопок инструкции
def create_instruction_button(script_name, button_name, col):
    """Создает кнопку инструкции в указанной колонке"""
    instruction_key = f"inst_{script_name}"
    
    # Кнопка инструкции (маленькая, с эмодзи)
    if col.button("ℹ️", key=instruction_key, help=f"Инструкция для {button_name}"):
        instruction_text = read_instruction_file(script_name)
        
        # Создаем модальное окно с инструкцией
        with st.expander(f"📖 Инструкция: {button_name}", expanded=True):
            st.markdown(instruction_text)
            
            # Кнопка для закрытия
            if st.button("✖️ Закрыть", key=f"close_{instruction_key}"):
                st.rerun()

# Функция для запуска скрипта
def run_script(script_path):
    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            encoding=SYSTEM_ENCODING,
            errors='replace',
            timeout=300
        )
        return {
            'success': result.returncode == 0,
            'output': result.stdout,
            'error': result.stderr
        }
    except subprocess.TimeoutExpired:
        return {
            'success': False,
            'output': '',
            'error': f'Таймаут: скрипт выполнялся более 5 минут'
        }
    except Exception as e:
        return {
            'success': False,
            'output': '',
            'error': f'Ошибка запуска: {str(e)}'
        }

# Создаем вкладки для двух групп скриптов
tab1, tab2 = st.tabs(["📋 Основные скрипты", "🔧 Дополнительные скрипты"])

# === Вкладка 1: Основные скрипты ===
with tab1:
    # Проверяем какие основные скрипты существуют
    available_main_scripts = {}
    missing_main_scripts = []

    for script_file, button_name in MAIN_SCRIPTS_CONFIG.items():
        if os.path.exists(script_file):
            available_main_scripts[script_file] = button_name
        else:
            missing_main_scripts.append(script_file)

    # Показываем доступные основные скрипты
    if available_main_scripts:
        st.subheader("✅ Доступные основные скрипты:")
        
        script_items = list(available_main_scripts.items())
        
        for script_file, button_name in script_items:
            # Получаем таймаут для скрипта
            cooldown = SCRIPT_TIMEOUTS.get(script_file, DEFAULT_TIMEOUT)
            button_key = f"main_{script_file}"
            
            # Проверяем, на таймауте ли кнопка
            on_cooldown, remaining = is_button_on_cooldown(button_key, cooldown)
            
            # Создаем строку с двумя колонками: для кнопки запуска и для кнопки инструкции
            col1, col2 = st.columns([5, 1])
            
            with col1:
                # Создаем текст для кнопки
                if on_cooldown:
                    button_text = f"⏳ {button_name} (подождите {remaining}с)"
                    disabled = True
                else:
                    button_text = button_name
                    disabled = False
                
                # Кнопка запуска
                if st.button(button_text, key=button_key, use_container_width=True, disabled=disabled):
                    # Обновляем время последнего нажатия
                    update_last_click(button_key)
                    
                    with st.spinner(f"Запускаю {button_name}..."):
                        result = run_script(script_file)
                        
                        if result['success']:
                            st.success(f"✅ {button_name} выполнен!")
                            if result['output']:
                                with st.expander("📋 Показать вывод", expanded=True):
                                    st.code(result['output'])
                        else:
                            st.error(f"❌ Ошибка в {button_name}")
                            if result['error']:
                                with st.expander("🔍 Показать ошибку", expanded=True):
                                    st.code(result['error'])
            
            # Кнопка инструкции в отдельной колонке
            create_instruction_button(script_file, button_name, col2)
            
            # Добавляем небольшой отступ между строками
            st.markdown("")

        # Кнопка для запуска всех основных скриптов
        st.markdown("---")
        
        # Для кнопки "запустить все" тоже добавим таймаут
        all_button_key = "run_all_main"
        all_cooldown = 10  # 10 секунд таймаут для запуска всех скриптов
        on_cooldown_all, remaining_all = is_button_on_cooldown(all_button_key, all_cooldown)
        
        if on_cooldown_all:
            all_button_text = f"⏳ ЗАПУСТИТЬ ВСЕ ОСНОВНЫЕ СКРИПТЫ (подождите {remaining_all}с)"
            disabled_all = True
        else:
            all_button_text = "⚡ ЗАПУСТИТЬ ВСЕ ОСНОВНЫЕ СКРИПТЫ"
            disabled_all = False
        
        if st.button(all_button_text, key=all_button_key, use_container_width=True, type="secondary", disabled=disabled_all):
            update_last_click(all_button_key)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_area = st.empty()
            
            all_output = "📊 РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ ОСНОВНЫХ СКРИПТОВ:\n\n"
            
            for i, (script_file, button_name) in enumerate(script_items):
                progress = (i) / len(script_items)
                progress_bar.progress(progress)
                status_text.text(f"🔄 Выполняется: {button_name}")
                
                result = run_script(script_file)
                
                all_output += f"=== {button_name} ===\n"
                if result['success']:
                    all_output += "✅ УСПЕХ\n"
                    if result['output']:
                        all_output += result['output'] + "\n"
                else:
                    all_output += "❌ ОШИБКА\n"
                    if result['error']:
                        all_output += result['error'] + "\n"
                all_output += "\n"
            
            progress_bar.progress(1.0)
            status_text.text("✅ Все основные скрипты выполнены!")
            
            with results_area.container():
                st.code(all_output)

# === Вкладка 2: Дополнительные скрипты ===
with tab2:
    # Проверяем какие дополнительные скрипты существуют
    available_extra_scripts = {}
    missing_extra_scripts = []

    for script_file, button_name in EXTRA_SCRIPTS_CONFIG.items():
        if os.path.exists(script_file):
            available_extra_scripts[script_file] = button_name
        else:
            missing_extra_scripts.append(script_file)

    # Показываем доступные дополнительные скрипты
    if available_extra_scripts:
        st.subheader("✅ Доступные дополнительные скрипты:")
        
        script_items = list(available_extra_scripts.items())
        
        for script_file, button_name in script_items:
            # Получаем таймаут для скрипта
            cooldown = SCRIPT_TIMEOUTS.get(script_file, DEFAULT_TIMEOUT)
            button_key = f"extra_{script_file}"
            
            # Проверяем, на таймауте ли кнопка
            on_cooldown, remaining = is_button_on_cooldown(button_key, cooldown)
            
            # Создаем строку с двумя колонками: для кнопки запуска и для кнопки инструкции
            col1, col2 = st.columns([5, 1])
            
            with col1:
                # Создаем текст для кнопки
                if on_cooldown:
                    button_text = f"⏳ {button_name} (подождите {remaining}с)"
                    disabled = True
                else:
                    button_text = button_name
                    disabled = False
                
                # Кнопка запуска
                if st.button(button_text, key=button_key, use_container_width=True, disabled=disabled):
                    # Обновляем время последнего нажатия
                    update_last_click(button_key)
                    
                    with st.spinner(f"Запускаю {button_name}..."):
                        result = run_script(script_file)
                        
                        if result['success']:
                            st.success(f"✅ {button_name} выполнен!")
                            if result['output']:
                                with st.expander("📋 Показать вывод", expanded=True):
                                    st.code(result['output'])
                        else:
                            st.error(f"❌ Ошибка в {button_name}")
                            if result['error']:
                                with st.expander("🔍 Показать ошибку", expanded=True):
                                    st.code(result['error'])
            
            # Кнопка инструкции в отдельной колонке
            create_instruction_button(script_file, button_name, col2)
            
            # Добавляем небольшой отступ между строками
            st.markdown("")

        # Кнопка для запуска всех дополнительных скриптов
        st.markdown("---")
        
        # Для кнопки "запустить все" тоже добавим таймаут
        all_button_key = "run_all_extra"
        all_cooldown = 10  # 10 секунд таймаут для запуска всех скриптов
        on_cooldown_all, remaining_all = is_button_on_cooldown(all_button_key, all_cooldown)
        
        if on_cooldown_all:
            all_button_text = f"⏳ ЗАПУСТИТЬ ВСЕ ДОПОЛНИТЕЛЬНЫЕ СКРИПТЫ (подождите {remaining_all}с)"
            disabled_all = True
        else:
            all_button_text = "⚡ ЗАПУСТИТЬ ВСЕ ДОПОЛНИТЕЛЬНЫЕ СКРИПТЫ"
            disabled_all = False
        
        if st.button(all_button_text, key=all_button_key, use_container_width=True, type="secondary", disabled=disabled_all):
            update_last_click(all_button_key)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_area = st.empty()
            
            all_output = "📊 РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ ДОПОЛНИТЕЛЬНЫХ СКРИПТОВ:\n\n"
            
            for i, (script_file, button_name) in enumerate(script_items):
                progress = (i) / len(script_items)
                progress_bar.progress(progress)
                status_text.text(f"🔄 Выполняется: {button_name}")
                
                result = run_script(script_file)
                
                all_output += f"=== {button_name} ===\n"
                if result['success']:
                    all_output += "✅ УСПЕХ\n"
                    if result['output']:
                        all_output += result['output'] + "\n"
                else:
                    all_output += "❌ ОШИБКА\n"
                    if result['error']:
                        all_output += result['error'] + "\n"
                all_output += "\n"
            
            progress_bar.progress(1.0)
            status_text.text("✅ Все дополнительные скрипты выполнены!")
            
            with results_area.container():
                st.code(all_output)

# Инструкция для пользователей
with st.sidebar:
    st.header("📖 Инструкция")
    
    st.markdown("""
    Здесь будет ссылка на confluence
    """)
    
    st.markdown("---")
    st.caption("🔄 Для обновления списка скриптов перезапусти приложение")









