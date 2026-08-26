import sqlite3
import hashlib
import secrets
from datetime import datetime, timedelta
import json

class Database:
    def __init__(self, db_path="data/teacher_assistant.db"):
        self.db_path = db_path
        self.init_database()
    
    def get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def init_database(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Таблица пользователей
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                assistant_name TEXT DEFAULT 'Помощник',
                user_nickname TEXT DEFAULT '',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_login TIMESTAMP,
                settings TEXT DEFAULT '{}'
            )
        ''')
        
        # Таблица заметок
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                content TEXT,
                reminder_date TIMESTAMP,
                is_completed BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                priority TEXT DEFAULT 'normal',
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Таблица истории чата
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS chat_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                message TEXT NOT NULL,
                response TEXT NOT NULL,
                mode TEXT DEFAULT 'text',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Таблица КТП
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS lesson_plans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                subject TEXT NOT NULL,
                class_name TEXT NOT NULL,
                quarter INTEGER NOT NULL,
                lesson_data TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Таблица списков классов
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS class_lists (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                class_name TEXT NOT NULL,
                student_data TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def register_user(self, username, password, assistant_name="Помощник", user_nickname=""):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        salt = secrets.token_hex(16)
        password_hash = hashlib.sha256((password + salt).encode()).hexdigest()
        
        try:
            cursor.execute('''
                INSERT INTO users (username, password_hash, salt, assistant_name, user_nickname)
                VALUES (?, ?, ?, ?, ?)
            ''', (username, password_hash, salt, assistant_name, user_nickname))
            
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()
    
    def authenticate_user(self, username, password):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        conn.close()
        
        if user:
            password_hash = hashlib.sha256((password + user['salt']).encode()).hexdigest()
            if password_hash == user['password_hash']:
                # Обновляем время последнего входа
                conn = self.get_connection()
                conn.execute('UPDATE users SET last_login = ? WHERE id = ?', 
                           (datetime.now(), user['id']))
                conn.commit()
                conn.close()
                return dict(user)
        
        return None
    
    def get_user_notes(self, user_id):
        conn = self.get_connection()
        notes = conn.execute(
            'SELECT * FROM notes WHERE user_id = ? ORDER BY created_at DESC',
            (user_id,)
        ).fetchall()
        conn.close()
        return [dict(note) for note in notes]
    
    def add_note(self, user_id, title, content, reminder_date=None, priority='normal'):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO notes (user_id, title, content, reminder_date, priority)
            VALUES (?, ?, ?, ?, ?)
        ''', (user_id, title, content, reminder_date, priority))
        conn.commit()
        note_id = cursor.lastrowid
        conn.close()
        return note_id
    
    def update_note(self, note_id, **kwargs):
        conn = self.get_connection()
        for key, value in kwargs.items():
            conn.execute(f'UPDATE notes SET {key} = ? WHERE id = ?', (value, note_id))
        conn.commit()
        conn.close()
    
    def delete_note(self, note_id):
        conn = self.get_connection()
        conn.execute('DELETE FROM notes WHERE id = ?', (note_id,))
        conn.commit()
        conn.close()
    
    def save_chat_message(self, user_id, message, response, mode='text'):
        conn = self.get_connection()
        conn.execute('''
            INSERT INTO chat_history (user_id, message, response, mode)
            VALUES (?, ?, ?, ?)
        ''', (user_id, message, response, mode))
        conn.commit()
        conn.close()
    
    def get_chat_history(self, user_id, limit=50):
        conn = self.get_connection()
        messages = conn.execute(
            'SELECT * FROM chat_history WHERE user_id = ? ORDER BY created_at DESC LIMIT ?',
            (user_id, limit)
        ).fetchall()
        conn.close()
        return [dict(msg) for msg in reversed(messages)]
    
    def save_class_list(self, user_id, class_name, student_data):
        conn = self.get_connection()
        conn.execute('''
            INSERT INTO class_lists (user_id, class_name, student_data)
            VALUES (?, ?, ?)
        ''', (user_id, class_name, json.dumps(student_data, ensure_ascii=False)))
        conn.commit()
        conn.close()
    
    def get_class_list(self, user_id, class_name):
        conn = self.get_connection()
        result = conn.execute(
            'SELECT * FROM class_lists WHERE user_id = ? AND class_name = ? ORDER BY created_at DESC LIMIT 1',
            (user_id, class_name)
        ).fetchone()
        conn.close()
        return dict(result) if result else None




    # Добавьте эти методы в класс Database
    
    def get_notes_by_date_range(self, user_id, start_date, end_date):
        """Получить заметки за период"""
        conn = self.get_connection()
        notes = conn.execute('''
            SELECT * FROM notes 
            WHERE user_id = ? AND reminder_date BETWEEN ? AND ?
            ORDER BY reminder_date, priority
        ''', (user_id, start_date, end_date)).fetchall()
        conn.close()
        return [dict(note) for note in notes]
    
    def get_notes_by_date(self, user_id, target_date):
        """Получить заметки на конкретную дату"""
        conn = self.get_connection()
        notes = conn.execute('''
            SELECT * FROM notes 
            WHERE user_id = ? AND DATE(reminder_date) = DATE(?)
            ORDER BY priority, created_at
        ''', (user_id, target_date)).fetchall()
        conn.close()
        return [dict(note) for note in notes]
    
    def get_all_notes_with_dates(self, user_id):
        """Получить все заметки с датами"""
        conn = self.get_connection()
        notes = conn.execute('''
            SELECT * FROM notes 
            WHERE user_id = ? AND reminder_date IS NOT NULL
            ORDER BY reminder_date
        ''', (user_id,)).fetchall()
        conn.close()
        return [dict(note) for note in notes]
    
    def get_upcoming_reminders(self, user_id, days=7):
        """Получить ближайшие напоминания"""
        conn = self.get_connection()
        today = datetime.now().strftime('%Y-%m-%d')
        end_date = (datetime.now() + timedelta(days=days)).strftime('%Y-%m-%d')
        
        notes = conn.execute('''
            SELECT * FROM notes 
            WHERE user_id = ? AND reminder_date BETWEEN ? AND ?
            AND is_completed = FALSE
            ORDER BY reminder_date, priority
        ''', (user_id, today, end_date)).fetchall()
        conn.close()
        return [dict(note) for note in notes]