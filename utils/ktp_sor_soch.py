from docx import Document
from io import BytesIO
import pandas as pd
import re
from datetime import date
from docx.oxml.ns import qn


def parse_dates(date_string):
    results = []
    if not date_string or date_string.strip() == '-':
        return results
    normalized = date_string.replace(';', '\n')
    pattern = r'(\d+[а-яА-ЯёЁ]*)\s*[-–—]\s*([\d\.\-]+)'
    matches = re.findall(pattern, normalized)
    for match in matches:
        class_name = match[0].strip()
        date_val = match[1].strip()
        results.append((class_name, date_val if date_val != '-' else '-'))
    return results


def determine_quarter(date_str):
    if not date_str or date_str == '-':
        return 'Не определена'
    try:
        parts = date_str.split('.')
        if len(parts) >= 2:
            day = int(parts[0])
            month = int(parts[1])
            year = 2026 if month >= 9 else 2027
            d = date(year, month, day)
            if date(2026, 9, 1) <= d <= date(2026, 10, 25):
                return '1 четверть'
            elif date(2026, 11, 2) <= d <= date(2026, 12, 29):
                return '2 четверть'
            elif date(2027, 1, 11) <= d <= date(2027, 3, 21):
                return '3 четверть'
            elif date(2027, 3, 29) <= d <= date(2027, 5, 25):
                return '4 четверть'
    except:
        pass
    return 'Не определена'


def find_sor_soch_in_text(text):
    matches = []
    seen = set()
    if not text:
        return matches
    
    text_lower = text.lower()
    
    # СОР ПЕРВЫМ
    sor_patterns = [
        r'сор\s*№\s*(\d+)',
        r'сор\s+(\d+)',
        r'сор№\s*(\d+)',
        r'сор(\d+)',
    ]
    
    for pattern in sor_patterns:
        for m in re.finditer(pattern, text_lower):
            num = m.group(1)
            key = ('СОР', num)
            if key not in seen:
                seen.add(key)
                matches.append({'type': 'СОР', 'number': num})
    
    # СОЧ ПОСЛЕ
    soch_patterns = [
        r'соч\s*№\s*(\d+)',
        r'соч\s+(\d+)',
        r'соч№\s*(\d+)',
        r'соч(\d+)',
    ]
    
    for pattern in soch_patterns:
        for m in re.finditer(pattern, text_lower):
            num = m.group(1)
            key = ('СОЧ', num)
            if key not in seen:
                seen.add(key)
                matches.append({'type': 'СОЧ', 'number': num})
    
    return matches


def check_horizontal_merge(row):
    try:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                for child in tcPr:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag == 'gridSpan':
                        return True
        return False
    except:
        return False


def read_ktp_and_extract_sor_soch(file_bytes, file_extension='docx'):
    try:
        doc = Document(BytesIO(file_bytes))
        assessments = []
        
        for table in doc.tables:
            if len(table.columns) < 6:
                continue
            
            for row_idx in range(1, len(table.rows)):
                row = table.rows[row_idx]
                
                if check_horizontal_merge(row):
                    continue
                
                try:
                    topic_text = row.cells[2].text.strip() if len(row.cells) > 2 else ""
                    date_text = row.cells[5].text.strip() if len(row.cells) > 5 else ""
                    note_text = row.cells[6].text.strip() if len(row.cells) > 6 else ""
                except:
                    continue
                
                matches = find_sor_soch_in_text(topic_text + " " + note_text)
                
                if not matches:
                    continue
                
                class_dates = parse_dates(date_text)
                
                for match in matches:
                    if class_dates:
                        for class_name, date_val in class_dates:
                            quarter = determine_quarter(date_val)
                            assessments.append({
                                'Класс': class_name,
                                'Четверть': quarter,
                                'Тип': match['type'],
                                'Номер': match['number'],
                                'Дата': date_val if date_val else '-'
                            })
                    else:
                        quarter = determine_quarter(date_text)
                        assessments.append({
                            'Класс': 'Не указан',
                            'Четверть': quarter,
                            'Тип': match['type'],
                            'Номер': match['number'],
                            'Дата': date_text if date_text else '-'
                        })
        
        if not assessments:
            return None, "СОР/СОЧ не найдены", {}
        
        def sort_key(item):
            cls = item['Класс']
            date_str = item['Дата']
            try:
                parts = date_str.split('.')
                if len(parts) >= 2:
                    d = int(parts[0])
                    m = int(parts[1])
                    y = 2026 if m >= 9 else 2027
                    return (cls, y, m, d)
            except:
                pass
            return (cls, 9999, 99, 99)
        
        assessments.sort(key=sort_key)
        df = pd.DataFrame(assessments)
        return df, None, {'rows_with_sor_soch': len(df)}
    
    except Exception as e:
        return None, f"Ошибка: {str(e)}", {}


def create_excel_schedule(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Общий график', index=False)
        for class_name in sorted(df['Класс'].unique()):
            if class_name == 'Не указан':
                continue
            class_df = df[df['Класс'] == class_name]
            sheet_name = f"Класс_{class_name}"[:31]
            class_df.to_excel(writer, sheet_name=sheet_name, index=False)
    buffer.seek(0)
    return buffer