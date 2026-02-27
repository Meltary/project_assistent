"""Модуль для работы с базой данных соответствий между Лоцман и 1С."""
import sqlite3
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Set, List, Tuple


# Путь к файлу базы данных
DB_PATH = Path(__file__).parent / "mappings.db"


def get_connection() -> sqlite3.Connection:
    """
    Создает и возвращает соединение с базой данных.
    
    Returns:
        Соединение с БД
    """
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row  # Для доступа к колонкам по имени
    return conn


def init_database() -> None:
    """
    Инициализирует базу данных, создавая необходимые таблицы.
    """
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name_loc TEXT NOT NULL,
                loc_path TEXT NOT NULL,
                name_1c TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(name_loc, loc_path, name_1c)
            )
        """)
        
        # Создаем индексы для быстрого поиска
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_name_loc_path 
            ON mappings(name_loc, loc_path)
        """)
        
        conn.commit()
    finally:
        conn.close()


def save_mappings(mappings: Dict[Tuple[str, str], Set[str]]) -> int:
    """
    Сохраняет соответствия в базу данных.
    
    Args:
        mappings: Словарь соответствий {(name_loc, loc_path): {name_1c, ...}}
    
    Returns:
        Количество сохраненных записей
    """
    init_database()
    conn = get_connection()
    count = 0
    
    try:
        cursor = conn.cursor()
        current_time = datetime.now().isoformat()
        
        for (name_loc, loc_path), name_1c_set in mappings.items():
            # Удаляем старые соответствия для этой пары (name_loc, loc_path)
            cursor.execute("""
                DELETE FROM mappings 
                WHERE name_loc = ? AND loc_path = ?
            """, (name_loc, loc_path))
            
            # Добавляем новые соответствия
            for name_1c in name_1c_set:
                try:
                    cursor.execute("""
                        INSERT INTO mappings (name_loc, loc_path, name_1c, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?)
                    """, (name_loc, loc_path, name_1c, current_time, current_time))
                    count += 1
                except sqlite3.IntegrityError:
                    # Если запись уже существует, обновляем updated_at
                    cursor.execute("""
                        UPDATE mappings 
                        SET updated_at = ? 
                        WHERE name_loc = ? AND loc_path = ? AND name_1c = ?
                    """, (current_time, name_loc, loc_path, name_1c))
        
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()
    
    return count


def load_mappings() -> Dict[Tuple[str, str], Set[str]]:
    """
    Загружает все соответствия из базы данных.
    
    Returns:
        Словарь соответствий {(name_loc, loc_path): {name_1c, ...}}
    """
    if not DB_PATH.exists():
        return {}
    
    conn = get_connection()
    mappings: Dict[Tuple[str, str], Set[str]] = {}
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT name_loc, loc_path, name_1c 
            FROM mappings
            ORDER BY updated_at DESC
        """)
        
        for row in cursor.fetchall():
            key = (row['name_loc'], row['loc_path'])
            if key not in mappings:
                mappings[key] = set()
            mappings[key].add(row['name_1c'])
    finally:
        conn.close()
    
    return mappings


def get_mappings_for_locman(name_loc: str, loc_path: str) -> Set[str]:
    """
    Получает соответствия для конкретного элемента Лоцман.
    
    Args:
        name_loc: Наименование из Лоцман
        loc_path: Путь к файлу Лоцман
    
    Returns:
        Множество наименований из 1С
    """
    if not DB_PATH.exists():
        return set()
    
    conn = get_connection()
    result = set()
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT name_1c 
            FROM mappings
            WHERE name_loc = ? AND loc_path = ?
        """, (name_loc, loc_path))
        
        for row in cursor.fetchall():
            result.add(row['name_1c'])
    finally:
        conn.close()
    
    return result


def delete_mappings(name_loc: Optional[str] = None, loc_path: Optional[str] = None) -> int:
    """
    Удаляет соответствия из базы данных.
    
    Args:
        name_loc: Наименование из Лоцман (опционально)
        loc_path: Путь к файлу Лоцман (опционально)
    
    Returns:
        Количество удаленных записей
    """
    if not DB_PATH.exists():
        return 0
    
    conn = get_connection()
    count = 0
    
    try:
        cursor = conn.cursor()
        if name_loc and loc_path:
            cursor.execute("""
                DELETE FROM mappings 
                WHERE name_loc = ? AND loc_path = ?
            """, (name_loc, loc_path))
        else:
            cursor.execute("DELETE FROM mappings")
        
        count = cursor.rowcount
        conn.commit()
    finally:
        conn.close()
    
    return count


def get_all_mappings_count() -> int:
    """
    Возвращает общее количество сохраненных соответствий.
    
    Returns:
        Количество записей в БД
    """
    if not DB_PATH.exists():
        return 0
    
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) as count FROM mappings")
        result = cursor.fetchone()
        return result['count'] if result else 0
    finally:
        conn.close()
