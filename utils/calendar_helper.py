from datetime import datetime, timedelta, date
import pandas as pd
import json


class KazakhstanSchoolCalendar:
    """Календарь учителя Республики Казахстан на 2026-2027 учебный год"""
    
    def __init__(self):
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month
        
        # ===== ПРАЗДНИКИ =====
        self.holidays_2026_2027 = {
            date(2026, 10, 25): {"name": "День Республики (воскресенье)", "is_holiday": True},
            date(2026, 10, 26): {"name": "Перенос: День Республики", "is_holiday": True},
            date(2026, 12, 16): {"name": "День Независимости (среда)", "is_holiday": True},
            date(2027, 1, 1): {"name": "Новый год (пятница)", "is_holiday": True},
            date(2027, 1, 2): {"name": "Новый год (суббота)", "is_holiday": True},
            date(2027, 1, 4): {"name": "Перенос: Новый год", "is_holiday": True},
            date(2027, 1, 7): {"name": "Рождество Христово (четверг)", "is_holiday": True},
            date(2027, 3, 8): {"name": "Международный женский день (понедельник)", "is_holiday": True},
            date(2027, 3, 15): {"name": "День Конституции (понедельник)", "is_holiday": True},
            date(2027, 3, 21): {"name": "Наурыз мейрамы (воскресенье)", "is_holiday": True},
            date(2027, 3, 22): {"name": "Наурыз мейрамы (понедельник)", "is_holiday": True},
            date(2027, 3, 23): {"name": "Наурыз мейрамы (вторник)", "is_holiday": True},
            date(2027, 3, 24): {"name": "Перенос: Наурыз мейрамы", "is_holiday": True},
            date(2027, 5, 1): {"name": "Праздник единства народа (суббота)", "is_holiday": True},
            date(2027, 5, 3): {"name": "Перенос: Праздник единства", "is_holiday": True},
            date(2027, 5, 7): {"name": "День защитника Отечества (пятница)", "is_holiday": True},
            date(2027, 5, 9): {"name": "День Победы (воскресенье)", "is_holiday": True},
            date(2027, 5, 10): {"name": "Перенос: День Победы", "is_holiday": True},
            date(2027, 5, 16): {"name": "Курбан айт (воскресенье)", "is_holiday": True},
        }
        
        # ===== ТОЧНЫЕ ДАТЫ ЧЕТВЕРТЕЙ 2026-2027 =====
        # Формат: (начало, конец, каникулы_начало, каникулы_конец)
        self.school_quarters = {
            2026: {
                1: (date(2026, 9, 1), date(2026, 10, 25), date(2026, 10, 26), date(2026, 11, 1)),
                2: (date(2026, 11, 2), date(2026, 12, 29), date(2026, 12, 30), date(2027, 1, 10)),
                3: (date(2027, 1, 11), date(2027, 3, 21), date(2027, 3, 22), date(2027, 3, 28)),
                4: (date(2027, 3, 29), date(2027, 5, 25), None, None),
            }
        }
    
    def get_current_school_year(self):
        return 2026
    
    def get_current_quarter(self):
        today = date.today()
        
        for quarter, (start, end, hol_start, hol_end) in self.school_quarters[2026].items():
            if start <= today <= end:
                return quarter
            if hol_start and hol_end and hol_start <= today <= hol_end:
                return f"каникулы_{quarter}"
        
        if today.month in [6, 7, 8]:
            return 0
        
        return None
    
    def get_quarter_dates(self, quarter, school_year=None):
        if school_year is None:
            school_year = 2026
        
        if school_year in self.school_quarters:
            quarter_data = self.school_quarters[school_year].get(quarter)
            if quarter_data:
                return quarter_data[0], quarter_data[1]
        
        return None, None
    
    def get_quarter_holidays(self, quarter, school_year=None):
        if school_year is None:
            school_year = 2026
        
        if school_year in self.school_quarters:
            quarter_data = self.school_quarters[school_year].get(quarter)
            if quarter_data and len(quarter_data) >= 4:
                return quarter_data[2], quarter_data[3]
        
        return None, None
    
    def is_holiday(self, check_date):
        if check_date in self.holidays_2026_2027:
            info = self.holidays_2026_2027[check_date]
            return True, info['name']
        return False, None
    
    def is_working_day(self, check_date):
        is_holiday, _ = self.is_holiday(check_date)
        if is_holiday:
            return False
        if check_date.weekday() >= 5:
            return False
        return True
    
    def get_working_days(self, start_date, end_date):
        working_days = []
        current = start_date
        while current <= end_date:
            if self.is_working_day(current):
                working_days.append(current)
            current += timedelta(days=1)
        return working_days
    
    def get_holidays_between(self, start_date, end_date):
        holidays = []
        current = start_date
        while current <= end_date:
            is_holiday, name = self.is_holiday(current)
            if is_holiday:
                holidays.append({
                    'date': current,
                    'name': name,
                    'weekday': current.strftime('%A')
                })
            current += timedelta(days=1)
        return holidays
    
    def get_days_until(self, target_date):
        today = date.today()
        return (target_date - today).days
    
    def get_working_days_between(self, start_date, end_date):
        if start_date > end_date:
            return 0
        return len(self.get_working_days(start_date, end_date))
    
    def get_days_until_end(self, end_date):
        today = date.today()
        if end_date < today:
            return 0
        return (end_date - today).days
    
    def get_weeks_until_end(self, end_date):
        days = self.get_days_until_end(end_date)
        return days // 7 if days > 0 else 0
    
    def get_dates_by_weekdays(self, start_date, end_date, weekdays, include_holidays=False):
        dates_by_day = {}
        day_names_ru = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']
        
        current = start_date
        while current <= end_date:
            if current.weekday() in weekdays:
                if include_holidays or self.is_working_day(current):
                    day_name = day_names_ru[current.weekday()]
                    if day_name not in dates_by_day:
                        dates_by_day[day_name] = []
                    
                    is_holiday, holiday_name = self.is_holiday(current)
                    date_info = current.strftime('%d.%m.%Y')
                    if is_holiday:
                        date_info += f" ({holiday_name})"
                    
                    dates_by_day[day_name].append(date_info)
            current += timedelta(days=1)
        
        return dates_by_day
    
    def format_dates_table(self, dates_by_day):
        if not dates_by_day:
            return pd.DataFrame()
        
        max_len = max(len(dates) for dates in dates_by_day.values())
        df_data = {}
        
        for day, dates in dates_by_day.items():
            df_data[day] = dates + [''] * (max_len - len(dates))
        
        return pd.DataFrame(df_data)
    
    def get_quarter_info(self, quarter, school_year=None):
        if school_year is None:
            school_year = 2026
        
        start, end = self.get_quarter_dates(quarter, school_year)
        hol_start, hol_end = self.get_quarter_holidays(quarter, school_year)
        
        if not start or not end:
            return None
        
        today = date.today()
        
        days_until_start = self.get_days_until(start)
        days_until_end = self.get_days_until(end)
        
        total_working_days = self.get_working_days_between(start, end)
        
        if today < start:
            status = "Четверть еще не началась"
            working_days_left = total_working_days
            days_passed = 0
        elif start <= today <= end:
            status = "Четверть идет"
            working_days_left = self.get_working_days_between(today, end)
            days_passed = (today - start).days
        else:
            status = "Четверть завершена"
            working_days_left = 0
            days_passed = (end - start).days + 1
        
        holidays = self.get_holidays_between(start, end)
        
        return {
            'school_year': school_year,
            'quarter': quarter,
            'start_date': start,
            'end_date': end,
            'holiday_start': hol_start,
            'holiday_end': hol_end,
            'total_days': (end - start).days + 1,
            'days_passed': days_passed,
            'working_days_count': total_working_days,
            'holidays': holidays,
            'status': status,
            'days_until_start': days_until_start,
            'days_until_end': days_until_end,
            'working_days_left': working_days_left,
            'weeks_until_start': days_until_start // 7 if days_until_start > 0 else 0,
            'weeks_until_end': days_until_end // 7 if days_until_end > 0 else 0
        }
    
    def get_all_quarters_info(self):
        quarters = []
        for q in range(1, 5):
            info = self.get_quarter_info(q)
            if info:
                quarters.append(info)
        return quarters
    
    def get_next_holidays(self, count=5):
        today = date.today()
        upcoming = []
        
        for holiday_date, info in self.holidays_2026_2027.items():
            if holiday_date >= today:
                upcoming.append({
                    'date': holiday_date,
                    'name': info['name'],
                    'days_left': (holiday_date - today).days,
                    'is_vacation': False
                })
        
        for quarter, (_, _, hol_start, hol_end) in self.school_quarters.get(2026, {}).items():
            if hol_start and hol_end and hol_start >= today:
                upcoming.append({
                    'date': hol_start,
                    'name': f"Каникулы после {quarter}-й четверти",
                    'days_left': (hol_start - today).days,
                    'is_vacation': True
                })
        
        upcoming.sort(key=lambda x: x['date'])
        return upcoming[:count]


class CalendarHelper(KazakhstanSchoolCalendar):
    """Совместимый класс"""
    
    def __init__(self):
        super().__init__()
    
    def get_quarter_dates(self, quarter_number, school_year=None):
        return super().get_quarter_dates(quarter_number, school_year)
    
    def get_working_days(self, start_date, end_date):
        return super().get_working_days(start_date, end_date)
    
    def get_days_until_end(self, end_date):
        return super().get_days_until_end(end_date)
    
    def get_weeks_until_end(self, end_date):
        return super().get_weeks_until_end(end_date)
    
    def get_dates_by_weekdays(self, start_date, end_date, weekdays, include_holidays=False):
        return super().get_dates_by_weekdays(start_date, end_date, weekdays, include_holidays)
    
    def format_dates_table(self, dates_by_day):
        return super().format_dates_table(dates_by_day)
    
    def get_quarter_info(self, quarter, school_year=None):
        return super().get_quarter_info(quarter, school_year)
    
    def get_current_quarter(self):
        return super().get_current_quarter()
    
    def get_current_school_year(self):
        return super().get_current_school_year()
    
    def get_all_quarters_info(self):
        return super().get_all_quarters_info()