from datetime import datetime, timedelta, date
import calendar
import pandas as pd


class NotesCalendar:
    """Календарь для заметок"""

    def __init__(self):
        self.month_names_ru = [
            '', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
            'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
        ]
        self.day_names_ru = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
        self.day_names_full = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

    def get_month_calendar(self, year, month):
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdayscalendar(year, month)
        return month_days

    def get_week_dates(self, target_date):
        monday = target_date - timedelta(days=target_date.weekday())
        week_dates = []
        for i in range(7):
            week_dates.append(monday + timedelta(days=i))
        return week_dates

    def get_week_range(self, target_date):
        week_dates = self.get_week_dates(target_date)
        return week_dates[0], week_dates[6]

    def format_week_label(self, target_date):
        start, end = self.get_week_range(target_date)
        return f"{start.strftime('%d.%m')} - {end.strftime('%d.%m.%Y')}"

    def is_today(self, check_date):
        if isinstance(check_date, date):
            return check_date == date.today()
        elif isinstance(check_date, datetime):
            return check_date.date() == date.today()
        else:
            return False

    def is_weekend(self, check_date):
        if isinstance(check_date, datetime):
            return check_date.weekday() >= 5
        elif isinstance(check_date, date):
            return check_date.weekday() >= 5
        else:
            return False

    def get_priority_color(self, priority):
        colors = {
            'high': '#FFD6D6',
            'normal': '#FFE4C2',
            'low': '#D9F2DF'
        }
        return colors.get(priority, '#FFFFFF')

    def get_priority_emoji(self, priority):
        emojis = {
            'high': '🔴',
            'normal': '🟠',
            'low': '🟢'
        }
        return emojis.get(priority, '⚪')

    def get_holiday_color(self):
        return '#FFF1B8'