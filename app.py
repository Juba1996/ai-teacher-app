import io
import json
import urllib.parse
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import google.generativeai as genai
from PIL import Image
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestClassifier

# ==========================================
# 0. СЛОВАРЬ ПЕРЕВОДОВ ИНТЕРФЕЙСА
# ==========================================
translations = {
    "ru": {
        "page_title": "Bilim AI — Помощник Учителя",
        "sidebar_title": "🎓 Bilim AI Platform",
        "api_subheader": "🔑 Доступ к ИИ",
        "api_help": "Введите ключ один раз для всех инструментов",
        "api_expander": "ℹ️ Как получить API ключ бесплатно?",
        "tools_subheader": "🛠️ Модули системы",
        "footer": "✨ Разработано для преподавателей",
        "menu": [
            "📝 Генератор карточек",
            "📅 AI-Генератор КТП",
            "📋 AI-Конструктор КСП",
            "📊 Анализ и визуализация (EDA)",
            "🤖 ML-Прогноз уровня ученика",
            "📷 AI-Проверка по фото",
            "👤 Генератор характеристик",
            "⚡ Разминки и интерактивы",
        ],
        "no_key": "Ключ не найден! Введите свой Gemini API Key слева или убедитесь, что настроен базовый ключ.",
        "warning_default_key": "⚠️ Вы используете общий API-ключ. При высокой нагрузке от других учителей он может временно не работать. Рекомендуем получить свой бесплатный ключ!",
        "ai_lang_prompt": "Напиши ответ строго на русском языке.",
        
        "source": "Источник данных:",
        "source_options": ["Google Таблица", "Excel-файл", "Сгенерировать через ИИ ✨"],
        "template_info": "💡 **Шаблон таблицы:** Ваша таблица должна содержать колонки: `Вопрос`, `Ответ`, `Сложность` (Легкий/Средний/Сложный). На листе учеников: `ФИО`.",
        "sheet_link": "Ссылка на Google Таблицу:",
        "upload_excel": "Загрузите Excel-файл (листы: Банк_вопросов, Ученики):",
        "success_excel": "Данные успешно прочитаны!",
        "ai_topic_lbl": "Тема для генерации вопросов:",
        "ai_students_lbl": "Список учеников (через запятую или с новой строки):",
        "settings": "Параметры генерации вариантов",
        "easy": "Легких вопросов:",
        "med": "Средних вопросов:",
        "hard": "Сложных вопросов:",
        "gen_word": "Сгенерировать варианты в Word",
        "wait_ai": "ИИ обрабатывает данные и формирует варианты...",
        "student_lbl": "Ученик(ца):",
        "task_lbl": "Задание",
        "answer_lbl": "Ответ: ____________________",
        "keys_title": "КЛЮЧИ (ДЛЯ УЧИТЕЛЯ)",
        "done": "Документы успешно созданы!",
        "download_cards": "Скачать Карточки (Word)",
        "download_keys": "Скачать Ключи (Word)",
        
        "subject": "Учебный предмет:",
        "grade": "Класс:",
        "quarters": "Количество четвертей:",
        "hours": "Часов в неделю:",
        "total_lessons": "Всего академических часов:",
        "source_pdf_text": "Источник тем:",
        "pdf_opt": ["Ввести темы текстом", "Загрузить PDF-файл"],
        "topics_lbl": "Перечень тем для распределения:",
        "gen_ktp": "Сгенерировать КТП в Word",
        "wait_ktp": "ИИ анализирует материалы и аккуратно собирает КТП...",
        "download_ktp": "Скачать КТП (Word)",
        
        "teacher_name": "ФИО преподавателя:",
        "topic_lbl": "Тема урока:",
        "target_lbl": "Цели обучения (ЦО):",
        "gen_ksp": "Сгенерировать КСП в Word",
        "wait_ksp": "ИИ методист разрабатывает структуру и этапы урока...",
        "download_ksp": "Скачать КСП (Word)",
        
        # ... остальные переводы EDA, ML, Фото, Характеристика (сокращено для экономии места, они остаются теми же)
        "eda_sub": "Анализ успеваемости класса",
        "eda_load": "Загрузить оценки (.xlsx)",
        "eda_select": "Показатель:",
        "eda_btn": "Анализ",
        "eda_wait": "Расчет...",
        
        "ml_sub": "Прогноз успеваемости",
        "ml_txt": "Введите данные:",
        "att": "Посещаемость (%):",
        "hw": "Домашка (%):",
        "test": "Тесты:",
        "activity": "Активность:",
        "act_opts": ["Низкая", "Средняя", "Высокая"],
        "ml_btn": "Прогноз",
        "ml_wait": "Анализ...",
        "rec": "Прогноз:",
        
        "photo_load": "Загрузите фото:",
        "photo_check": "Проверить",
        "photo_wait": "Распознавание...",
        
        "char_sub": "Генератор характеристик",
        "name_lbl": "ФИО:",
        "cls_lbl": "Класс:",
        "att_lbl": "Посещаемость:",
        "perf_lbl": "Успеваемость:",
        "perf_opts": ["Отличник", "Ударник", "Слабо"],
        "beh_lbl": "Дисциплина:",
        "beh_opts": ["Хорошая", "Средняя", "Плохая"],
        "traits_lbl": "Доп:",
        "char_btn": "Создать",
        "char_wait": "Пишу...",
        
        "warm_sub": "Разминки",
        "warm_top": "Тема:",
        "warm_time": "Мин:",
        "warm_btn": "Найти",
        "warm_wait": "Ищу..."
    },
    "kk": {
        "page_title": "Bilim AI — Мұғалім Көмекшісі",
        "sidebar_title": "🎓 Bilim AI Platform",
        "api_subheader": "🔑 ЖИ қолжетімділік кілті",
        "api_help": "Барлық құралдар үшін кілтті бір рет енгізіңіз",
        "api_expander": "ℹ️ API кілтін қалай алуға болады?",
        "tools_subheader": "🛠️ Жүйе модульдері",
        "footer": "✨ Оқытушылар үшін әзірленген",
        "menu": [
            "📝 Тапсырма карточкаларын жасау",
            "📅 КТП AI-Генераторы",
            "📋 ҚМЖ (КСП) AI-Конструкторы",
            "📊 Талдау және визуализация (EDA)",
            "🤖 Оқушы деңгейін ML болжау",
            "📷 Фото арқылы AI тексеру",
            "👤 Мінездеме генераторы",
            "⚡ Сергіту сәттері мен интерактив",
        ],
        "no_key": "Кілт табылмады! Өз Gemini API кілтіңізді енгізіңіз.",
        "warning_default_key": "⚠️ Сіз жалпы API кілтін пайдаланып жатырсыз. Жүктеме көп болғанда істемей қалуы мүмкін. Өз кілтіңізді алуға кеңес береміз!",
        "ai_lang_prompt": "Жауапты қатаң түрде қазақ тілінде жаз.",
        
        "source": "Дереккөз:",
        "source_options": ["Google Кесте", "Excel-файл", "ЖИ арқылы генерациялау ✨"],
        "template_info": "💡 **Кесте шаблоны:** Сіздің кестеңізде мына бағандар болуы керек: `Сұрақ`, `Жауап`, `Қиындығы` (Жеңіл/Орташа/Қиын). Оқушылар парағында: `Аты-жөні`.",
        "sheet_link": "Google кестенің сілтемесі:",
        "upload_excel": "Excel файлын жүктеңіз (Банк_вопросов, Ученики):",
        "success_excel": "Деректер сәтті оқылды!",
        "ai_topic_lbl": "Сұрақтар құруға арналған тақырып:",
        "ai_students_lbl": "Оқушылар тізімі (үтір арқылы немесе жаңа жолдан):",
        "settings": "Генерация параметрлері",
        "easy": "Жеңіл сұрақтар:",
        "med": "Орташа сұрақтар:",
        "hard": "Қиын сұрақтар:",
        "gen_word": "Word форматында нұсқалар жасау",
        "wait_ai": "ЖИ деректерді өңдеп, нұсқаларды жасауда...",
        "student_lbl": "Оқушы:",
        "task_lbl": "Тапсырма",
        "answer_lbl": "Жауап: ____________________",
        "keys_title": "ЖАУАПТАР (МҰҒАЛІМГЕ)",
        "done": "Құжаттар сәтті дайындалды!",
        "download_cards": "Карточкаларды жүктеу (Word)",
        "download_keys": "Жауаптарды жүктеу (Word)",
        
        "subject": "Оқу пәні:",
        "grade": "Сынып / Курс:",
        "quarters": "Тоқсан саны:",
        "hours": "Аптасына сағат:",
        "total_lessons": "Барлық сағат:",
        "source_pdf_text": "Тақырыптар көзі:",
        "pdf_opt": ["Тақырыптарды мәтінмен енгізу", "PDF файлын жүктеу"],
        "topics_lbl": "Тақырыптар тізімі:",
        "gen_ktp": "Word форматында КТП құру",
        "wait_ktp": "ЖИ КТП жасауда...",
        "download_ktp": "КТП жүктеу (Word)",
        
        "teacher_name": "Оқытушының А.Т.Ә.:",
        "topic_lbl": "Сабақ тақырыбы:",
        "target_lbl": "Оқыту мақсаттары (ОМ):",
        "gen_ksp": "Word форматында ҚМЖ құру",
        "wait_ksp": "ЖИ сабақ жоспарын әзірлеуде...",
        "download_ksp": "ҚМЖ жүктеу (Word)",
        
        "eda_sub": "Анализ", "eda_load": "Жүктеу", "eda_select": "Көрсеткіш:", "eda_btn": "Анализ", "eda_wait": "Күте тұрыңыз...",
        "ml_sub": "Болжам", "ml_txt": "Деректер:", "att": "Қатысу:", "hw": "Үй жұмысы:", "test": "Тест:", "activity": "Белсенділік:", "act_opts": ["Төмен", "Орташа", "Жоғары"], "ml_btn": "Болжау", "ml_wait": "Күте тұрыңыз...", "rec": "Болжам:",
        "photo_load": "Фото:", "photo_check": "Тексеру", "photo_wait": "Күте тұрыңыз...",
        "char_sub": "Мінездеме", "name_lbl": "Аты-жөні:", "cls_lbl": "Сынып:", "att_lbl": "Қатысу:", "perf_lbl": "Үлгерім:", "perf_opts": ["Үздік", "Екпінді", "Орташа"], "beh_lbl": "Тәртіп:", "beh_opts": ["Жақсы", "Орташа", "Нашар"], "traits_lbl": "Қосымша:", "char_btn": "Құру", "char_wait": "Күте тұрыңыз...",
        "warm_sub": "Сергіту", "warm_top": "Тақырып:", "warm_time": "Уақыт:", "warm_btn": "Құру", "warm_wait": "Күте тұрыңыз..."
    }
}

# ==========================================
# 1. НАСТРОЙКИ СТРАНИЦЫ И СТИЛИ
# ==========================================
st.set_page_config(page_title="Bilim AI", page_icon="🎓", layout="wide")
st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {font-family: 'Plus Jakarta Sans', sans-serif;}
    .stApp {background: linear-gradient(135deg, #f4f6f9 0%, #edf2f7 100%);}
    .block-container {background-color: #ffffff; border-radius: 24px; padding: 3rem; box-shadow: 0 10px 30px rgba(0,0,0,0.04); margin-top: 2rem; margin-bottom: 2rem;}
    [data-testid="stSidebar"] {background-color: #0f172a; color: #ffffff;}
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] label {color: #e2e8f0 !important;}
    [data-testid="stSidebar"] .stRadio label p {color: #f8fafc !important; font-weight: 500;}
    .stButton>button {border-radius: 12px; background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%); color: white !important; font-weight: 600;}
</style>""", unsafe_allow_html=True)

# ⚠️ ВАШ ДЕФОЛТНЫЙ КЛЮЧ (ВСТАВИТЬ СЮДА) ⚠️
DEFAULT_API_KEY = ""

# ==========================================
# 2. БОКОВОЕ МЕНЮ И УПРАВЛЕНИЕ КЛЮЧАМИ
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/1972/1972413.png", width=60)
lang_choice = st.sidebar.selectbox("🌐 Тіл / Язык интерфейса:", ["Русский", "Қазақша"], index=0)
lang = "ru" if lang_choice == "Русский" else "kk"
t = translations[lang]

st.sidebar.markdown(f"### {t['sidebar_title']}")
st.sidebar.divider()
st.sidebar.subheader(t["api_subheader"])
user_api_key = st.sidebar.text_input("Gemini API Key:", type="password", help=t["api_help"])

with st.sidebar.expander(t["api_expander"]):
    st.markdown("1. Зайдите на [Google AI Studio](https://aistudio.google.com/app/apikey).\n2. Нажмите **Create API key**.\n3. Вставьте ключ выше.")

# Логика подхвата ключа и предупреждения
active_key = user_api_key.strip()
if not active_key:
    if DEFAULT_API_KEY:
        active_key = DEFAULT_API_KEY
        st.sidebar.warning(t["warning_default_key"], icon="⚠️")

st.sidebar.divider()
menu_choice = st.sidebar.radio("Navigation:", t["menu"], label_visibility="collapsed")


# ==========================================
# ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ОЧИСТКИ JSON ОТ ИИ
# ==========================================
def clean_json_response(text):
    text = text.strip()
    match = re.search(r'\[.*\]', text, re.DOTALL) if '[' in text else re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        return json.loads(match.group(0))
    return json.loads(text)


# ==========================================
# МОДУЛЬ 1: ГЕНЕРАТОР КАРТОЧЕК
# ==========================================
if menu_choice in ["📝 Генератор карточек", "📝 Тапсырма карточкаларын жасау"]:
    st.title(menu_choice)
    st.info(t["template_info"])
    st.divider()

    source_type = st.radio(t["source"], t["source_options"], horizontal=True)
    df_questions, df_students = None, None
    students_list = []

    if "Google" in source_type:
        sheet_url = st.text_input(f"🔗 {t['sheet_link']}", value="https://docs.google.com/spreadsheets/d/1fJKlRP7YY3r6DFjd_PuLXFIKkg3GdSRAM9Rxwq502e8/edit?usp=sharing")
        if sheet_url and "/d/" in sheet_url:
            sheet_id = sheet_url.split("/d/")[1].split("/")[0]
            try:
                df_questions = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=Банк_вопросов")
                df_students = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=Ученики")
                students_list = df_students.iloc[:, 0].dropna().tolist()
            except: pass
            
    elif "Excel" in source_type:
        uploaded_excel = st.file_uploader(f"📂 {t['upload_excel']}", type=["xlsx"])
        if uploaded_excel:
            xls = pd.ExcelFile(uploaded_excel)
            df_questions = pd.read_excel(xls, 'Банк_вопросов')
            df_students = pd.read_excel(xls, 'Ученики')
            students_list = df_students.iloc[:, 0].dropna().tolist()
            st.success(t["success_excel"])
            
    else: # ГЕНЕРАЦИЯ ЧЕРЕЗ ИИ
        ai_topic = st.text_input(f"🧠 {t['ai_topic_lbl']}", "Устройство компьютера и память")
        ai_students_raw = st.text_area(f"👥 {t['ai_students_lbl']}", "Иванов Иван\nПетров Петр\nСмирнова Анна")
        students_list = [s.strip() for s in ai_students_raw.replace(',', '\n').split('\n') if s.strip()]

    st.markdown(f"#### ⚙️ {t['settings']}")
    col1, col2, col3 = st.columns(3)
    with col1: count_easy = st.number_input(f"🟢 {t['easy']}", min_value=0, max_value=5, value=1)
    with col2: count_med = st.number_input(f"🟡 {t['med']}", min_value=0, max_value=5, value=1)
    with col3: count_hard = st.number_input(f"🔴 {t['hard']}", min_value=0, max_value=5, value=1)

    if st.button(f"🚀 {t['gen_word']}", type="primary", use_container_width=True):
        if not active_key: st.error(t["no_key"])
        elif not students_list: st.warning("Добавьте учеников!")
        else:
            with st.spinner(f"⏳ {t['wait_ai']}"):
                try:
                    # Если выбран ИИ, сначала генерируем базу вопросов
                    if "ИИ" in source_type or "ЖИ" in source_type:
                        genai.configure(api_key=active_key)
                        model = genai.GenerativeModel("gemini-3.6-flash")
                        total_q = (count_easy + count_med + count_hard) * 3 # Генерируем с запасом
                        prompt = f"{t['ai_lang_prompt']} Сгенерируй базу из {count_easy*3} легких, {count_med*3} средних и {count_hard*3} сложных вопросов по теме '{ai_topic}'. Верни строго JSON массив: [{{'вопрос': '...', 'ответ': '...', 'сложность': 'Легкий'}}, ...]"
                        res = model.generate_content(prompt)
                        q_data = clean_json_response(res.text)
                        df_questions = pd.DataFrame(q_data)

                    # Стандартизация колонок базы
                    df_questions.columns = df_questions.columns.astype(str).str.strip().str.lower()
                    rename_dict = {}
                    for col in df_questions.columns:
                        if "сложн" in col or "күрдел" in col or "қиын" in col: rename_dict[col] = "сложность"
                        elif "вопрос" in col or "сұрақ" in col: rename_dict[col] = "вопрос"
                        elif "ответ" in col or "жауап" in col: rename_dict[col] = "ответ"
                    df_questions = df_questions.rename(columns=rename_dict)

                    doc_students = Document()
                    doc_teacher = Document()
                    doc_teacher.add_heading(t["keys_title"], level=1)
                    
                    structure = {"Легкий": count_easy, "Средний": count_med, "Сложный": count_hard}
                    # Для казахского языка поддержка перевода сложности
                    if lang == "kk": structure = {"Жеңіл": count_easy, "Орташа": count_med, "Қиын": count_hard}

                    for student in students_list:
                        variant_questions = []
                        for level, count in structure.items():
                            if count > 0:
                                subset = df_questions[df_questions["сложность"].astype(str).str.strip().str.capitalize().str.contains(level[:3], case=False, na=False)]
                                if len(subset) == 0: subset = df_questions # Если уровень не найден, берем любые
                                variant_questions.append(subset.sample(n=min(count, len(subset))))
                        
                        student_variant = pd.concat(variant_questions).reset_index(drop=True)
                        
                        # Блок ученика
                        title = doc_students.add_heading("Проверочная работа" if lang=="ru" else "Бақылау жұмысы", level=2)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc_students.add_paragraph().add_run(f"{t['student_lbl']} {student}").bold = True
                        
                        # Блок учителя
                        doc_teacher.add_paragraph().add_run(f"\n👤 {student}").bold = True

                        for idx, row in student_variant.iterrows():
                            # В документ ученика
                            p_q = doc_students.add_paragraph()
                            p_q.add_run(f"{t['task_lbl']} {idx + 1}. ").bold = True
                            p_q.add_run(f"{row.get('вопрос', 'Ошибка вопроса')}\n")
                            p_q.add_run(t["answer_lbl"])
                            
                            # В документ учителя
                            doc_teacher.add_paragraph(f"  • {t['task_lbl']} {idx + 1}: {row.get('ответ', 'Нет ответа')}")
                            
                        doc_students.add_paragraph("--------------------------------------------------")

                    bio_students, bio_teacher = io.BytesIO(), io.BytesIO()
                    doc_students.save(bio_students)
                    doc_teacher.save(bio_teacher)
                    
                    st.success(f"🎉 {t['done']}")
                    col_d1, col_d2 = st.columns(2)
                    with col_d1: st.download_button(f"📄 {t['download_cards']}", bio_students.getvalue(), "Карточки.docx", use_container_width=True)
                    with col_d2: st.download_button(f"🔑 {t['download_keys']}", bio_teacher.getvalue(), "Ключи.docx", use_container_width=True)
                except Exception as e: st.error(f"Произошла ошибка при обработке данных: {e}")

# ==========================================
# МОДУЛЬ 2: AI-ГЕНЕРАТОР КТП (ИСПРАВЛЕНЫ ТАБЛИЦЫ)
# ==========================================
elif menu_choice in ["📅 AI-Генератор КТП", "📅 КТП AI-Генераторы"]:
    st.title(menu_choice)
    st.divider()
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        subject = st.text_input(t["subject"], "Информатика")
        grade = st.number_input(t["grade"], 1, 11, 8)
    with col_p2:
        quarters_count = st.selectbox(t["quarters"], [1, 2, 3, 4], index=0)
        hours_per_week = st.number_input(t["hours"], 1, 5, 2)
    
    quarters_weeks = {q: st.number_input(f"{q}-я четверть (недель):" if lang=="ru" else f"{q}-ші тоқсан (апта):", 1, 15, 8) for q in range(1, quarters_count + 1)}
    total_all_lessons = sum(q_w * hours_per_week for q_w in quarters_weeks.values())
    st.info(f"💡 {t['total_lessons']} **{total_all_lessons}**")

    textbook_content = st.text_area(t["topics_lbl"], "1. Алгоритмы\n2. Циклы Python\n3. Базы данных", height=100)

    if st.button(f"🚀 {t['gen_ktp']}", type="primary", use_container_width=True):
        if not active_key: st.error(t["no_key"])
        else:
            try:
                with st.spinner(f"⏳ {t['wait_ktp']}"):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"{t['ai_lang_prompt']} Составь КТП по предмету {subject}, {grade} класс, уроков: {total_all_lessons}. Темы: {textbook_content}. Верни строго JSON массив (БЕЗ markdown): [{{\"quarter\":1, \"lesson_num\":1, \"topic\":\"...\", \"targets\":\"...\", \"homework\":\"...\"}}]"
                    res = model.generate_content(prompt)
                    ktp_data = clean_json_response(res.text)

                    doc = Document()
                    section = doc.sections[-1]
                    section.orientation = WD_ORIENT.LANDSCAPE
                    section.page_width, section.page_height = section.page_height, section.page_width

                    title = doc.add_heading("КАЛЕНДАРНО-ТЕМАТИЧЕСКОЕ ПЛАНИРОВАНИЕ (КТП)" if lang=="ru" else "КҮНТІЗБЕЛІК-ТАҚЫРЫПТЫҚ ЖОСПАР (КТП)", level=1)
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    table = doc.add_table(rows=1, cols=5)
                    table.style = 'Table Grid' # Обязательная сетка
                    
                    headers = ["Четверть", "№", "Тема урока", "Цели обучения", "Домашнее задание"]
                    widths = [0.8, 0.5, 3.5, 3.5, 1.5] # Задаем правильную ширину колонок в дюймах

                    hdr_cells = table.rows[0].cells
                    for i, h in enumerate(headers):
                        hdr_cells[i].text = h
                        hdr_cells[i].paragraphs[0].runs[0].bold = True # Жирный заголовок

                    for item in ktp_data:
                        row = table.add_row().cells
                        row[0].text = str(item.get("quarter", ""))
                        row[1].text = str(item.get("lesson_num", ""))
                        row[2].text = str(item.get("topic", ""))
                        row[3].text = str(item.get("targets", ""))
                        row[4].text = str(item.get("homework", ""))
                    
                    # Применяем ширину ко всем ячейкам таблицы
                    for row in table.rows:
                        for idx, width in enumerate(widths):
                            row.cells[idx].width = Inches(width)

                    bio = io.BytesIO()
                    doc.save(bio)
                    st.success(f"🎉 {t['done']}")
                    st.download_button(f"📄 {t['download_ktp']}", bio.getvalue(), f"КТП_{subject}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка парсинга или ИИ: {e}")

# ==========================================
# МОДУЛЬ 3: AI-КОНСТРУКТОР КСП (ИСПРАВЛЕНЫ ТАБЛИЦЫ)
# ==========================================
elif menu_choice in ["📋 AI-Конструктор КСП", "📋 ҚМЖ (КСП) AI-Конструкторы"]:
    st.title(menu_choice)
    st.divider()
    col_k1, col_k2 = st.columns(2)
    with col_k1:
        teacher_name = st.text_input(t["teacher_name"], "Иванов И.И.")
        subject_ksp = st.text_input(t["subject"], "Информатика")
        grade_ksp = st.number_input(t["grade"], 1, 11, 8)
    with col_k2:
        topic_ksp = st.text_input(t["topic_lbl"], "Условный оператор IF")
        target_ksp = st.text_input(t["target_lbl"], "Уметь писать ветвления на Python")

    if st.button(f"🚀 {t['gen_ksp']}", type="primary", use_container_width=True):
        if not active_key: st.error(t["no_key"])
        else:
            try:
                with st.spinner(f"⏳ {t['wait_ksp']}"):
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel("gemini-3.6-flash")
                    prompt = f"{t['ai_lang_prompt']} Создай план урока по предмету {subject_ksp}, тема {topic_ksp}. Верни строго JSON (БЕЗ markdown): {{\"lesson_targets\":\"...\", \"eval_criteria\":\"...\", \"stages\":[{{\"time\":\"Начало\", \"teacher\":\"...\", \"student\":\"...\", \"eval\":\"...\", \"resources\":\"...\"}}]}}"
                    res = model.generate_content(prompt)
                    ksp_data = clean_json_response(res.text)

                    doc = Document()
                    doc.add_heading("КРАТКОСРОЧНЫЙ ПЛАН УРОКА (ҚМЖ)", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER

                    t_table = doc.add_table(rows=7, cols=2)
                    t_table.style = 'Table Grid'
                    info = [("Учитель:", teacher_name), ("Предмет:", subject_ksp), ("Класс:", str(grade_ksp)), ("Тема:", topic_ksp), ("ЦО:", target_ksp), ("Цели:", ksp_data.get("lesson_targets","")), ("Критерии:", ksp_data.get("eval_criteria",""))]
                    for idx, (l, v) in enumerate(info):
                        t_table.rows[idx].cells[0].text = l
                        t_table.rows[idx].cells[0].paragraphs[0].runs[0].bold = True
                        t_table.rows[idx].cells[1].text = str(v)

                    doc.add_paragraph()
                    
                    s_table = doc.add_table(rows=1, cols=5)
                    s_table.style = 'Table Grid'
                    headers = ["Этап", "Действия учителя", "Действия ученика", "Оценивание", "Ресурсы"]
                    widths = [1.0, 2.5, 2.5, 1.5, 1.5]
                    
                    hdr_cells = s_table.rows[0].cells
                    for i, h in enumerate(headers):
                        hdr_cells[i].text = h
                        hdr_cells[i].paragraphs[0].runs[0].bold = True

                    for stg in ksp_data.get("stages", []):
                        row = s_table.add_row().cells
                        row[0].text = str(stg.get("time",""))
                        row[1].text = str(stg.get("teacher",""))
                        row[2].text = str(stg.get("student",""))
                        row[3].text = str(stg.get("eval",""))
                        row[4].text = str(stg.get("resources",""))
                        
                    for row in s_table.rows:
                        for idx, width in enumerate(widths):
                            row.cells[idx].width = Inches(width)

                    bio = io.BytesIO()
                    doc.save(bio)
                    st.success(f"🎉 {t['done']}")
                    st.download_button(f"📄 {t['download_ksp']}", bio.getvalue(), f"КСП_{topic_ksp}.docx", use_container_width=True)
            except Exception as e: st.error(f"Ошибка ИИ: {e}")

# ==========================================
# ОСТАЛЬНЫЕ МОДУЛИ (EDA, ML, ФОТО, ХАРАКТЕРИСТИКА) остаются без изменений
# ==========================================
elif menu_choice in ["📊 Анализ и визуализация (EDA)", "📊 Талдау және визуализация (EDA)"]:
    st.title(menu_choice)
    st.info("В разработке / Остается как в предыдущей версии")

# И так далее для остальных разделов... (я опустил их код здесь, чтобы не дублировать, вы можете просто добавить их из прошлой версии)
