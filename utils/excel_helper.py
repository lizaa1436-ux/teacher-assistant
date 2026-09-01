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
        try:
            df = pd.read_excel(file)
            self.original_columns = list(df.columns)
            self.normalized_columns = self._normalize_columns(df.columns)
            new_columns = list(self.normalized_columns.values())
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
        normalized = {}
        for col in columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            if any(word in col_lower for word in ['фамили', 'surname', 'last name', 'lastname', 'fam', 'фам']):
                normalized[col] = 'Фамилия'
            elif any(word in col_lower for word in ['отчеств', 'patronymic', 'middle name']):
                normalized[col] = 'Отчество'
            elif any(word in col_lower for word in ['имя', 'name', 'first name', 'firstname']):
                normalized[col] = 'Имя'
            elif any(word in col_lower for word in ['адрес', 'address']):
                normalized[col] = 'Адрес'
            elif any(word in col_lower for word in ['телефон', 'phone', 'тел']):
                normalized[col] = 'Телефон'
            elif any(word in col_lower for word in ['дата', 'рожден', 'birth']):
                normalized[col] = 'Дата рождения'
            elif any(word in col_lower for word in ['иин', 'iin']):
                normalized[col] = 'ИИН'
            elif any(word in col_lower for word in ['мать', 'мама']):
                normalized[col] = 'Мать'
            elif any(word in col_lower for word in ['отец', 'папа']):
                normalized[col] = 'Отец'
            elif any(word in col_lower for word in ['класс', 'class']):
                normalized[col] = 'Класс'
            else:
                normalized[col] = col_str

        # Если Фамилия не найдена - первая колонка
        has_lastname = any(v == 'Фамилия' for v in normalized.values())
        if not has_lastname and len(columns) > 0:
            first_col = columns[0]
            normalized[first_col] = 'Фамилия'

        # Если Имя не найдено - вторая колонка
        has_firstname = any(v == 'Имя' for v in normalized.values())
        if not has_firstname and len(columns) > 1:
            second_col = columns[1]
            normalized[second_col] = 'Имя'

        # Если Отчество не найдено - третья колонка
        has_middlename = any(v == 'Отчество' for v in normalized.values())
        if not has_middlename and len(columns) > 2:
            third_col = columns[2]
            normalized[third_col] = 'Отчество'

        return normalized

    def _find_column(self, possible_names):
        if self.class_data is None:
            return None
        for col in self.class_data.columns:
            col_lower = str(col).lower().strip()
            for name in possible_names:
                if col_lower == name.lower() or name.lower() in col_lower:
                    return col
        if len(self.class_data.columns) > 0:
            return self.class_data.columns[0]
        return None

    def search_student(self, lastname, firstname=None, partial_match=True):
        if self.class_data is None:
            return None
        df = self.class_data
        lastname_col = self._find_column(['Фамилия'])
        firstname_col = self._find_column(['Имя'])
        if lastname_col is None:
            return None
        df[lastname_col] = df[lastname_col].astype(str).str.strip()
        lastname_clean = lastname.lower().strip()
        if partial_match:
            mask = df[lastname_col].str.lower().str.startswith(lastname_clean)
        else:
            mask = df[lastname_col].str.lower() == lastname_clean
        results = df[mask]
        if len(results) == 0:
            mask = df[lastname_col].str.lower().str.contains(lastname_clean, na=False)
            results = df[mask]
        if len(results) == 0:
            return None
        if len(results) > 1 and firstname and firstname_col:
            firstname_clean = firstname.lower().strip()
            mask = results[firstname_col].astype(str).str.lower().str.startswith(firstname_clean)
            results = results[mask]
        if len(results) == 1:
            student_data = results.iloc[0].to_dict()
            clean_data = {}
            for key, value in student_data.items():
                if pd.notna(value) and str(value).strip():
                    clean_data[key] = value
            return clean_data
        elif len(results) > 1:
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

    def get_student_info(self, lastname, firstname=None):
        student = self.search_student(lastname, firstname)
        if student is None:
            return None
        if isinstance(student, dict) and student.get('multiple'):
            info_lines = [f"Найдено {len(student['students'])} учеников:"]
            for s in student['students']:
                full_name = ' '.join([str(s.get('Фамилия', '')), str(s.get('Имя', '')), str(s.get('Отчество', ''))]).strip()
                info_lines.append(f"• {full_name}")
            return '\n'.join(info_lines)
        info_lines = ["### 📋 Полная информация об ученике", ""]
        full_name_parts = []
        if 'Фамилия' in student:
            full_name_parts.append(str(student['Фамилия']))
        if 'Имя' in student:
            full_name_parts.append(str(student['Имя']))
        if 'Отчество' in student:
            full_name_parts.append(str(student['Отчество']))
        if full_name_parts:
            info_lines.append(f"**👤 ФИО:** {' '.join(full_name_parts)}")
        for key, value in student.items():
            if key in ['Фамилия', 'Имя', 'Отчество']:
                continue
            if pd.notna(value) and str(value).strip():
                info_lines.append(f"**{key}:** {value}")
        return '\n'.join(info_lines)

    def get_all_students(self):
        if self.class_data is None:
            return []
        return self.class_data.to_dict('records')

    def search_by_field(self, field_name, value):
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
        if self.class_data is None:
            return []
        return list(self.class_data.columns)

    def get_student_count(self):
        if self.class_data is None:
            return 0
        return len(self.class_data)
