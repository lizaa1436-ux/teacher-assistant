import pandas as pd
import re
import json
from datetime import datetime, date


class ExcelHelper:
    def __init__(self):
        self.class_data = None
        self.original_columns = []
        self.normalized_columns = {}
    
    def load_class_list(self, file):
        """
        Загрузить список класса из Excel (социальный паспорт).
        Удаляет дубликаты колонок.
        """
        try:
            df = pd.read_excel(file)
            
            self.original_columns = list(df.columns)
            self.normalized_columns = self._normalize_columns(df.columns)
            
            new_columns = list(self.normalized_columns.values())
            
            # Удаляем дубликаты
            seen = {}
            unique_columns = []
            for col in new_columns:
                col_str = str(col).strip()
                if col_str in seen:
                    seen[col_str] += 1
                    unique_columns.append(f"{col_str}_{seen[col_str]}")
                else:
                    seen[col_str] = 1
                    unique_columns.append(col_str)
            
            df.columns = unique_columns
            self.class_data = df
            
            return df, self.normalized_columns
        
        except Exception as e:
            raise Exception(f"Ошибка загрузки файла: {str(e)}")
    
    def _normalize_columns(self, columns):
        """
        Нормализовать названия колонок для социального паспорта.
        """
        normalized = {}
        
        for col in columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            
            # Фамилия
            if any(word in col_lower for word in ['фамили', 'surname', 'last name', 'lastname', 'fam']):
                normalized[col] = 'Фамилия'
            
            # Имя
            elif any(word in col_lower for word in ['имя', 'name', 'first name', 'firstname']):
                normalized[col] = 'Имя'
            
            # Отчество
            elif any(word in col_lower for word in ['отчеств', 'patronymic', 'middle name', 'middlename']):
                normalized[col] = 'Отчество'
            
            # Адрес
            elif any(word in col_lower for word in ['адрес', 'address', 'место жительства', 'прожива', 'домашний адрес']):
                normalized[col] = 'Адрес'
            
            # Телефон
            elif any(word in col_lower for word in ['телефон', 'phone', 'тел', 'мобиль', 'mobile', 'сот']):
                normalized[col] = 'Телефон'
            
            # Дата рождения
            elif any(word in col_lower for word in ['дата рождения', 'день рождения', 'birthday', 'birth', 'рожден']):
                normalized[col] = 'Дата рождения'
            
            # ИИН
            elif any(word in col_lower for word in ['иин', 'iin', 'индивидуальн']):
                normalized[col] = 'ИИН'
            
            # Мать (ФИО матери)
            elif 'мать' in col_lower or 'мама' in col_lower or 'mother' in col_lower:
                if 'тел' in col_lower or 'phone' in col_lower:
                    normalized[col] = 'Телефон матери'
                elif 'место работы' in col_lower or 'работ' in col_lower:
                    normalized[col] = 'Место работы матери'
                else:
                    normalized[col] = 'ФИО матери'
            
            # Отец (ФИО отца)
            elif 'отец' in col_lower or 'папа' in col_lower or 'father' in col_lower:
                if 'тел' in col_lower or 'phone' in col_lower:
                    normalized[col] = 'Телефон отца'
                elif 'место работы' in col_lower or 'работ' in col_lower:
                    normalized[col] = 'Место работы отца'
                else:
                    normalized[col] = 'ФИО отца'
            
            # Родители (общее)
            elif any(word in col_lower for word in ['родител', 'parent', 'законн']):
                normalized[col] = 'Родители'
            
            # Семья
            elif any(word in col_lower for word in ['семь', 'family']):
                if 'статус' in col_lower or 'категор' in col_lower:
                    normalized[col] = 'Статус семьи'
                elif 'многодетн' in col_lower:
                    normalized[col] = 'Многодетная семья'
                elif 'малообеспеч' in col_lower:
                    normalized[col] = 'Малообеспеченная семья'
                elif 'неполн' in col_lower:
                    normalized[col] = 'Неполная семья'
                else:
                    normalized[col] = 'Семья'
            
            # Социальный статус
            elif any(word in col_lower for word in ['соц', 'статус', 'категор', 'социальн']):
                normalized[col] = 'Социальный статус'
            
            # Национальность
            elif any(word in col_lower for word in ['национальн', 'nationality', 'этнос']):
                normalized[col] = 'Национальность'
            
            # Пол
            elif any(word in col_lower for word in ['пол', 'gender', 'sex']):
                normalized[col] = 'Пол'
            
            # Группа здоровья
            elif any(word in col_lower for word in ['здоров', 'health', 'физкультур']):
                normalized[col] = 'Группа здоровья'
            
            # Инвалидность
            elif any(word in col_lower for word in ['инвалид', 'disabilit', 'овз']):
                normalized[col] = 'Инвалидность/ОВЗ'
            
            # Опека
            elif any(word in col_lower for word in ['опек', 'guardian', 'попечител']):
                normalized[col] = 'Опека/Попечительство'
            
            # Питание
            elif any(word in col_lower for word in ['питани', 'бесплатн', 'food']):
                normalized[col] = 'Питание'
            
            # Проезд
            elif any(word in col_lower for word in ['проезд', 'подвоз', 'транспорт']):
                normalized[col] = 'Проезд/Подвоз'
            
            # Email
            elif any(word in col_lower for word in ['email', 'e-mail', 'почта', 'mail']):
                normalized[col] = 'Email'
            
            # Класс
            elif any(word in col_lower for word in ['класс', 'class']):
                normalized[col] = 'Класс'
            
            # Примечание
            elif any(word in col_lower for word in ['примечан', 'заметк', 'коммент', 'дополнит']):
                normalized[col] = 'Примечание'
            
            # Если не определили — оставляем как есть
            else:
                normalized[col] = col_str
        
        return normalized
    
    def search_student(self, lastname, firstname=None, partial_match=True):
        """
        Гибкий поиск ученика по фамилии.
        Возвращает ВСЕ данные ученика.
        """
        if self.class_data is None:
            return None
        
        df = self.class_data
        
        lastname_col = self._find_column(['Фамилия'])
        firstname_col = self._find_column(['Имя'])
        
        if lastname_col is None:
            lastname_col = df.columns[0] if len(df.columns) > 0 else None
        
        if lastname_col is None:
            return None
        
        # Очищаем от пробелов
        df[lastname_col] = df[lastname_col].astype(str).str.strip()
        
        # Поиск по фамилии
        lastname_clean = lastname.lower().strip()
        
        if partial_match:
            mask = df[lastname_col].str.lower().str.startswith(lastname_clean)
        else:
            mask = df[lastname_col].str.lower() == lastname_clean
        
        results = df[mask]
        
        if len(results) == 0:
            # Пробуем contains
            mask = df[lastname_col].str.lower().str.contains(lastname_clean, na=False)
            results = df[mask]
        
        if len(results) == 0:
            return None
        
        # Если несколько и указано имя
        if len(results) > 1 and firstname and firstname_col:
            firstname_clean = firstname.lower().strip()
            mask = results[firstname_col].astype(str).str.lower().str.startswith(firstname_clean)
            results = results[mask]
        
        if len(results) == 1:
            # Возвращаем ВСЕ данные ученика
            student_data = results.iloc[0].to_dict()
            # Убираем NaN значения
            clean_data = {}
            for key, value in student_data.items():
                if pd.notna(value) and str(value).strip():
                    clean_data[key] = value
            return clean_data
        
        elif len(results) > 1:
            # Несколько учеников
            students_list = []
            for _, row in results.iterrows():
                student_info = {}
                if lastname_col:
                    student_info['Фамилия'] = row.get(lastname_col, '')
                if firstname_col:
                    student_info['Имя'] = row.get(firstname_col, '')
                middlename_col = self._find_column(['Отчество'])
                if middlename_col:
                    student_info['Отчество'] = row.get(middlename_col, '')
                students_list.append(student_info)
            
            return {'multiple': True, 'students': students_list}
        
        return None
    
    def _find_column(self, possible_names):
        """Найти колонку по возможным названиям"""
        if self.class_data is None:
            return None
        
        for col in self.class_data.columns:
            col_lower = str(col).lower().strip()
            for name in possible_names:
                if col_lower == name.lower() or name.lower() in col_lower:
                    return col
        
        return None
    
    def get_student_info(self, lastname, firstname=None):
        """
        Получить полную информацию об ученике в красивом формате.
        """
        student = self.search_student(lastname, firstname)
        
        if student is None:
            return None
        
        if isinstance(student, dict) and student.get('multiple'):
            info_lines = [f"Найдено {len(student['students'])} учеников с фамилией '{lastname}':"]
            for s in student['students']:
                full_name = ' '.join([
                    str(s.get('Фамилия', '')),
                    str(s.get('Имя', '')),
                    str(s.get('Отчество', ''))
                ]).strip()
                info_lines.append(f"• {full_name}")
            return '\n'.join(info_lines)
        
        # Один ученик — выводим ВСЕ данные
        info_lines = ["### 📋 Полная информация об ученике", ""]
        
        # Сначала ФИО
        full_name_parts = []
        if 'Фамилия' in student:
            full_name_parts.append(str(student['Фамилия']))
        if 'Имя' in student:
            full_name_parts.append(str(student['Имя']))
        if 'Отчество' in student:
            full_name_parts.append(str(student['Отчество']))
        
        if full_name_parts:
            info_lines.append(f"**👤 ФИО:** {' '.join(full_name_parts)}")
        
        # Остальные данные
        for key, value in student.items():
            if key in ['Фамилия', 'Имя', 'Отчество']:
                continue
            
            if pd.notna(value) and str(value).strip():
                info_lines.append(f"**{key}:** {value}")
        
        return '\n'.join(info_lines)
    
    def get_all_students(self):
        """Получить список всех учеников"""
        if self.class_data is None:
            return []
        return self.class_data.to_dict('records')
    
    def search_by_field(self, field_name, value):
        """Поиск по любому полю"""
        if self.class_data is None:
            return None
        
        col = self._find_column([field_name])
        if col is None:
            return None
        
        mask = self.class_data[col].astype(str).str.lower().str.contains(value.lower(), na=False)
        results = self.class_data[mask]
        
        if len(results) == 0:
            return None
        elif len(results) == 1:
            return results.iloc[0].to_dict()
        else:
            return results.to_dict('records')
    
    def get_column_names(self):
        """Названия колонок"""
        if self.class_data is None:
            return []
        return list(self.class_data.columns)
    
    def get_student_count(self):
        """Количество учеников"""
        if self.class_data is None:
            return 0
        return len(self.class_data)
