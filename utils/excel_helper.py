import pandas as pd
import json

class ExcelHelper:
    def __init__(self):
        self.class_data = None
    
    def load_class_list(self, file):
        """Загрузить список класса из Excel"""
        try:
            df = pd.read_excel(file)
            required_columns = ['Фамилия', 'Имя', 'Отчество']
            
            # Проверяем наличие обязательных колонок
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Отсутствует колонка '{col}' в файле")
            
            self.class_data = df
            return df
        except Exception as e:
            raise Exception(f"Ошибка загрузки файла: {str(e)}")
    
    def search_student(self, lastname, firstname=None):
        """Поиск ученика по фамилии и имени"""
        if self.class_data is None:
            return None
        
        # Поиск по фамилии
        results = self.class_data[self.class_data['Фамилия'].str.lower() == lastname.lower()]
        
        if len(results) == 0:
            return None
        elif len(results) == 1:
            return results.iloc[0].to_dict()
        else:
            # Если несколько учеников с такой фамилией
            if firstname:
                results = results[results['Имя'].str.lower() == firstname.lower()]
                if len(results) == 1:
                    return results.iloc[0].to_dict()
            
            # Возвращаем список для уточнения
            return {
                'multiple': True,
                'students': results[['Фамилия', 'Имя', 'Отчество']].to_dict('records')
            }
    
    def get_student_info(self, lastname, firstname=None):
        """Получить всю информацию об ученике"""
        student = self.search_student(lastname, firstname)
        
        if student and not isinstance(student.get('multiple'), bool):
            info = []
            for key, value in student.items():
                if pd.notna(value):
                    info.append(f"**{key}:** {value}")
            return '\n'.join(info)
        
        return None