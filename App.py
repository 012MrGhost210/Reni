# streamlit_app.py

import streamlit as st
import subprocess
import sys
import os
import locale

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
    "simple.py": "📁 Скопировать файл Daily Income 2026"
}

# Дополнительная группа скриптов
EXTRA_SCRIPTS_CONFIG = {
    "zaprosstavok.py": "📋 Запрос ставок",
    "stop.py": "📈 Убрать заглушку"
}
# ===============================================================

# Определяем системную кодировку
try:
    SYSTEM_ENCODING = locale.getpreferredencoding()
except:
    SYSTEM_ENCODING = 'cp1251'

st.title("🚀 Запуск Python скриптов")
st.markdown("---")

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
    # Извлекаем эмодзи из названия кнопки для кнопки инструкции
    emoji = button_name.split()[0] if button_name else "📄"
    
    # Создаем уникальный ключ для кнопки инструкции
    instruction_key = f"inst_{script_name}"
    
    # Кнопка инструкции (маленькая, с эмодзи)
    if col.button("ℹ️", key=instruction_key, help=f"Инструкция для {button_name}"):
        instruction_text = read_instruction_file(script_name)
        
        # Создаем модальное окно с инструкцией
        with st.expander(f"📖 Инструкция: {button_name}", expanded=True):
            st.markdown(instruction_text)
            
            # Кнопка для закрытия (просто для удобства)
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
        
        # Создаем колонки для кнопок (теперь каждая кнопка запуска будет в своей колонке)
        script_items = list(available_main_scripts.items())
        
        for script_file, button_name in script_items:
            # Создаем строку с двумя колонками: для кнопки запуска и для кнопки инструкции
            col1, col2 = st.columns([5, 1])
            
            with col1:
                if st.button(button_name, key=f"main_{script_file}", use_container_width=True):
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
        
        if st.button("⚡ ЗАПУСТИТЬ ВСЕ ОСНОВНЫЕ СКРИПТЫ", use_container_width=True, type="secondary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_area = st.empty()
            
            all_output = "📊 РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ ОСНОВНЫХ СКРИПТОВ:\n\n"
            
            script_items = list(available_main_scripts.items())
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
            # Создаем строку с двумя колонками: для кнопки запуска и для кнопки инструкции
            col1, col2 = st.columns([5, 1])
            
            with col1:
                if st.button(button_name, key=f"extra_{script_file}", use_container_width=True):
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
        
        if st.button("⚡ ЗАПУСТИТЬ ВСЕ ДОПОЛНИТЕЛЬНЫЕ СКРИПТЫ", use_container_width=True, type="secondary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_area = st.empty()
            
            all_output = "📊 РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ ДОПОЛНИТЕЛЬНЫХ СКРИПТОВ:\n\n"
            
            script_items = list(available_extra_scripts.items())
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



