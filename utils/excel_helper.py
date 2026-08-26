import pandas as pd
import re
import json
from datetime import datetime


class ExcelHelper:
    def __init__(self):
        self.class_data = None
        self.original_columns = []
        self.normalized_columns = {}
    
    def load_class_list(self, file):
        """
        Загрузить список класса из Excel.
        Автоматически определяет и нормализует названия колонок.
        """
        try:
            # Читаем Excel
            df = pd.read_excel(file)
            
            # Сохраняем оригинальные названия
            self.original_columns = list(df.columns)
            
            # Нормализуем названия колонок
            self.normalized_columns = self._normalize_columns(df.columns)
            
            # Переименовываем колонки
            df.columns = list(self.normalized_columns.values())
            
            self.class_data = df
            
            return df, self.normalized_columns
        
        except Exception as e:
            raise Exception(f"Ошибка загрузки файла: {str(e)}")
    
    def _normalize_columns(self, columns):
        """
        Нормализовать названия колонок.
        Определяет стандартные поля: фамилия, имя, отчество, адрес, телефон и т.д.
        """
        normalized = {}
        
        for col in columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            
            # Фамилия
            if any(word in col_lower for word in ['фамили', 'surname', 'last name', 'lastname', 'fam']):
                normalized[col] = 'Фамилия'
            
            # Имя
            elif any(word in col_lower for word in ['имя', 'name', 'first name', 'firstname', 'first']):
                if 'отчество' not in col_lower and 'фамили' not in col_lower:
                    normalized[col] = 'Имя'
                else:
                    normalized[col] = col_str
            
            # Отчество
            elif any(word in col_lower for word in ['отчеств', 'patronymic', 'middle name', 'middlename', 'middle']):
                normalized[col] = 'Отчество'
            
            # Адрес
            elif any(word in col_lower for word in ['адрес', 'address', 'место жительства', 'прожива']):
                normalized[col] = 'Адрес'
            
            # Телефон
            elif any(word in col_lower for word in ['телефон', 'phone', 'тел', 'мобиль', 'mobile', 'сот']):
                normalized[col] = 'Телефон'
            
            # Дата рождения
            elif any(word in col_lower for word in ['дата рождения', 'день рождения', 'birthday', 'birth', 'дата рожд']):
                normalized[col] = 'Дата рождения'
            
            # Родители
            elif any(word in col_lower for word in ['мать', 'мама', 'mother']):
                normalized[col] = 'Мать'
            elif any(word in col_lower for word in ['отец', 'папа', 'father']):
                normalized[col] = 'Отец'
            elif any(word in col_lower for word in ['родител', 'parent']):
                normalized[col] = 'Родители'
            
            # Email
            elif any(word in col_lower for word in ['email', 'e-mail', 'почта', 'mail']):
                normalized[col] = 'Email'
            
            # Класс
            elif any(word in col_lower for word in ['класс', 'class']):
                normalized[col] = 'Класс'
            
            # Пол
            elif any(word in col_lower for word in ['пол', 'gender', 'sex']):
                normalized[col] = 'Пол'
            
            # Национальность
            elif any(word in col_lower for word in ['национальн', 'nationality']):
                normalized[col] = 'Национальность'
            
            # ИИН
            elif any(word in col_lower for word in ['иин', 'iin', 'индивидуаль']):
                normalized[col] = 'ИИН'
            
            # Если не определили — оставляем как есть
            else:
                normalized[col] = col_str
        
        return normalized
    
    def search_student(self, lastname, firstname=None, partial_match=True):
        """
        Гибкий поиск ученика по фамилии и имени.
        
        Параметры:
        - lastname: фамилия (обязательно)
        - firstname: имя (опционально)
        - partial_match: True - частичное совпадение, False - точное
        """
        if self.class_data is None:
            return None
        
        df = self.class_data
        
        # Находим колонку с фамилией
        lastname_col = self._find_column(['Фамилия', 'фамилия', 'surname', 'lastname'])
        firstname_col = self._find_column(['Имя', 'имя', 'name', 'firstname'])
        
        if lastname_col is None:
            # Пробуем первую колонку
            lastname_col = df.columns[0] if len(df.columns) > 0 else None
        
        if lastname_col is None:
            return None
        
        # Очищаем от лишних пробелов
        df[lastname_col] = df[lastname_col].astype(str).str.strip()
        
        # Поиск по фамилии
        if partial_match:
            # Частичное совпадение (начинается с...)
            mask = df[lastname_col].str.lower().str.startswith(lastname.lower().strip())
        else:
            # Точное совпадение
            mask = df[lastname_col].str.lower() == lastname.lower().strip()
        
        results = df[mask]
        
        if len(results) == 0:
            # Если не нашли — пробуем contains (содержит)
            mask = df[lastname_col].str.lower().str.contains(lastname.lower().strip(), na=False)
            results = df[mask]
        
        if len(results) == 0:
            return None
        
        # Если несколько результатов и указано имя
        if len(results) > 1 and firstname and firstname_col:
            firstname = firstname.lower().strip()
            
            if partial_match:
                mask = results[firstname_col].astype(str).str.lower().str.startswith(firstname)
            else:
                mask = results[firstname_col].astype(str).str.lower() == firstname
            
            results = results[mask]
        
        if len(results) == 1:
            return results.iloc[0].to_dict()
        elif len(results) > 1:
            # Несколько учеников
            students_list = []
            for _, row in results.iterrows():
                student_info = {}
                if lastname_col:
                    student_info['Фамилия'] = row.get(lastname_col, '')
                if firstname_col:
                    student_info['Имя'] = row.get(firstname_col, '')
                # Добавляем отчество если есть
                middlename_col = self._find_column(['Отчество', 'отчество'])
                if middlename_col:
                    student_info['Отчество'] = row.get(middlename_col, '')
                students_list.append(student_info)
            
            return {
                'multiple': True,
                'students': students_list
            }
        
        return None
    
    def _find_column(self, possible_names):
        """
        Найти колонку по возможным названиям
        """
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
        Получить полную информацию об ученике в красивом формате
        """
        student = self.search_student(lastname, firstname)
        
        if student is None:
            return None
        
        if isinstance(student, dict) and student.get('multiple'):
            # Несколько учеников
            info_lines = [f"Найдено {len(student['students'])} учеников с фамилией '{lastname}':"]
            for s in student['students']:
                full_name = ' '.join([
                    s.get('Фамилия', ''),
                    s.get('Имя', ''),
                    s.get('Отчество', '')
                ]).strip()
                info_lines.append(f"• {full_name}")
            return '\n'.join(info_lines)
        
        # Один ученик — выводим всю информацию
        info_lines = []
        for key, value in student.items():
            if pd.notna(value) and str(value).strip():
                info_lines.append(f"**{key}:** {value}")
        
        return '\n'.join(info_lines)
    
    def get_all_students(self):
        """Получить список всех учеников"""
        if self.class_data is None:
            return []
        return self.class_data.to_dict('records')
    
    def search_by_field(self, field_name, value):
        """
        Поиск по любому полю
        """
        if self.class_data is None:
            return None
        
        col = self._find_column([field_name])
        
        if col is None:
            return None
        
        df = self.class_data
        mask = df[col].astype(str).str.lower().str.contains(value.lower(), na=False)
        
        results = df[mask]
        
        if len(results) == 0:
            return None
        elif len(results) == 1:
            return results.iloc[0].to_dict()
        else:
            return results.to_dict('records')
    
    def get_unique_values(self, field_name):
        """Получить уникальные значения из колонки"""
        if self.class_data is None:
            return []
        
        col = self._find_column([field_name])
        if col is None:
            return []
        
        return self.class_data[col].dropna().unique().tolist()
    
    def get_column_names(self):
        """Получить список нормализованных названий колонок"""
        if self.class_data is None:
            return []
        return list(self.class_data.columns)
    
    def get_student_count(self):
        """Количество учеников"""
        if self.class_data is None:
            return 0
        return len(self.class_data)
