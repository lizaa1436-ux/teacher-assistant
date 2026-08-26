import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
import io
from datetime import date

# Настройка страницы
st.set_page_config(
    page_title="График СОР/СОЧ",
    page_icon="📊",
    layout="wide"
)

st.title("📊 График СОР и СОЧ")
st.write("Загрузите КТП для автоматического извлечения СОР/СОЧ")

uploaded_file = st.file_uploader("📁 Загрузите КТП", type=['docx'])

if uploaded_file:
    doc = Document(BytesIO(uploaded_file.read()))
    
    results = []
    debug_info = []
    
    for table in doc.tables:
        for row_idx in range(1, len(table.rows)):
            row = table.rows[row_idx]
            
            try:
                topic = row.cells[2].text.strip()
                dates = row.cells[5].text.strip()
                note = row.cells[6].text.strip() if len(row.cells) > 6 else ""
            except:
                continue
            
            full_text = topic + " " + note
            text_lower = full_text.lower()
            
            # ПОИСК СОР
            sor_nums = re.findall(r'сор[^\d]*(\d+)', text_lower)
            # ПОИСК СОЧ
            soch_nums = re.findall(r'соч[^\d]*(\d+)', text_lower)
            
            if sor_nums or soch_nums:
                debug_info.append(f"Строка {row_idx}: СОР={sor_nums}, СОЧ={soch_nums} | {topic[:40]}")
                
                # Разбираем даты
                for line in dates.split('\n'):
                    m = re.match(r'(\d+[а-яА-ЯёЁ]*)\s*-\s*([\d\.\-]+)', line.strip())
                    if m:
                        cls = m.group(1).strip()
                        dt = m.group(2).strip()
                        
                        # Определяем четверть
                        quarter = "Не определена"
                        if dt and dt != '-':
                            try:
                                parts = dt.split('.')
                                if len(parts) >= 2:
                                    day = int(parts[0])
                                    month = int(parts[1])
                                    d = date(2026 if month >= 9 else 2027, month, day)
                                    
                                    if date(2026, 9, 1) <= d <= date(2026, 10, 25):
                                        quarter = "1 четверть"
                                    elif date(2026, 11, 2) <= d <= date(2026, 12, 29):
                                        quarter = "2 четверть"
                                    elif date(2027, 1, 11) <= d <= date(2027, 3, 21):
                                        quarter = "3 четверть"
                                    elif date(2027, 3, 29) <= d <= date(2027, 5, 25):
                                        quarter = "4 четверть"
                            except:
                                pass
                        
                        for num in sor_nums:
                            results.append({
                                'Класс': cls,
                                'Четверть': quarter,
                                'Тип': 'СОР',
                                'Номер': num,
                                'Дата': dt if dt else '-'
                            })
                        
                        for num in soch_nums:
                            results.append({
                                'Класс': cls,
                                'Четверть': quarter,
                                'Тип': 'СОЧ',
                                'Номер': num,
                                'Дата': dt if dt else '-'
                            })
    
    # Показываем отладку
    with st.expander("🔍 Отладка", expanded=False):
        for line in debug_info:
            st.write(line)
    
    if results:
        # Сортировка
        def sort_key(item):
            cls = item['Класс']
            dt = item['Дата']
            try:
                parts = dt.split('.')
                if len(parts) >= 2:
                    d = int(parts[0])
                    m = int(parts[1])
                    y = 2026 if m >= 9 else 2027
                    return (cls, y, m, d)
            except:
                pass
            return (cls, 9999, 99, 99)
        
        results.sort(key=sort_key)
        df = pd.DataFrame(results)
        
        st.success(f"✅ Найдено {len(df)} записей")
        
        # Статистика
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("СОР", len(df[df['Тип'] == 'СОР']))
        with col2:
            st.metric("СОЧ", len(df[df['Тип'] == 'СОЧ']))
        with col3:
            st.metric("Классов", len(df[df['Класс'] != '-']['Класс'].unique()))
        
        # Таблица
        st.dataframe(df, use_container_width=True)
        
        # Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Общий график', index=False)
            for cls in sorted(df['Класс'].unique()):
                cls_df = df[df['Класс'] == cls]
                sheet_name = f"Класс_{cls}"[:31]
                cls_df.to_excel(writer, sheet_name=sheet_name, index=False)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Скачать Excel",
            data=buffer,
            file_name="график_СОР_СОЧ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.warning("СОР/СОЧ не найдены")