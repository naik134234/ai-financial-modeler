"""
Database module for AI Financial Modeler
Handles job persistence using SQLite
"""

import os
import sqlite3
from datetime import datetime
from typing import Optional, List, Dict, Any
import json
from contextlib import contextmanager

# Database path
import tempfile

# Database path
# Use /tmp on Vercel (read-only filesystem elsewhere)
if os.environ.get("VERCEL"):
    DB_PATH = os.path.join(tempfile.gettempdir(), "models.db")
else:
    DB_PATH = os.path.join(os.path.dirname(__file__), "models.db")


def get_connection():
    """Get database connection"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@contextmanager
def get_db():
    """Context manager for database connections"""
    conn = get_connection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """Initialize database tables"""
    with get_db() as conn:
        cursor = conn.cursor()
        
        # Jobs table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                company_name TEXT NOT NULL,
                symbol TEXT,
                industry TEXT,
                model_type TEXT DEFAULT 'general',
                status TEXT DEFAULT 'pending',
                progress INTEGER DEFAULT 0,
                message TEXT,
                file_path TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                completed_at TIMESTAMP,
                forecast_years INTEGER DEFAULT 5,
                source TEXT DEFAULT 'stock',
                request_data TEXT,
                result_data TEXT
            )
        """)
        
        # Model metrics table (for preview)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS model_metrics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id TEXT NOT NULL,
                metric_name TEXT NOT NULL,
                metric_value REAL,
                metric_format TEXT DEFAULT 'number',
                FOREIGN KEY (job_id) REFERENCES jobs(id)
            )
        """)
        
        # User preferences table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS preferences (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        
        conn.commit()


# Job CRUD operations

def create_job(
    job_id: str,
    company_name: str,
    symbol: Optional[str] = None,
    industry: Optional[str] = None,
    forecast_years: int = 5,
    source: str = "stock",
    request_data: Optional[Dict] = None
) -> Dict[str, Any]:
    """Create a new job"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO jobs (id, company_name, symbol, industry, forecast_years, source, request_data)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            job_id, 
            company_name, 
            symbol, 
            industry, 
            forecast_years, 
            source,
            json.dumps(request_data) if request_data else None
        ))
    return get_job(job_id)


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    """Get job by ID"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM jobs WHERE id = ?", (job_id,))
        row = cursor.fetchone()
        if row:
            return dict(row)
    return None


def update_job(job_id: str, **kwargs) -> Optional[Dict[str, Any]]:
    """Update job fields"""
    if not kwargs:
        return get_job(job_id)
    
    # Handle special fields
    if 'request_data' in kwargs and isinstance(kwargs['request_data'], dict):
        kwargs['request_data'] = json.dumps(kwargs['request_data'])
    if 'result_data' in kwargs and isinstance(kwargs['result_data'], dict):
        kwargs['result_data'] = json.dumps(kwargs['result_data'])
    
    set_clause = ", ".join(f"{k} = ?" for k in kwargs.keys())
    values = list(kwargs.values()) + [job_id]
    
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(f"UPDATE jobs SET {set_clause} WHERE id = ?", values)
    
    return get_job(job_id)


def get_job_history(limit: int = 50, offset: int = 0) -> List[Dict[str, Any]]:
    """Get job history ordered by creation date"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM jobs 
            ORDER BY created_at DESC 
            LIMIT ? OFFSET ?
        """, (limit, offset))
        return [dict(row) for row in cursor.fetchall()]


def get_completed_jobs(limit: int = 20) -> List[Dict[str, Any]]:
    """Get completed jobs for re-download"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM jobs 
            WHERE status = 'completed' AND file_path IS NOT NULL
            ORDER BY completed_at DESC 
            LIMIT ?
        """, (limit,))
        return [dict(row) for row in cursor.fetchall()]


def delete_job(job_id: str) -> bool:
    """Delete a job"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM model_metrics WHERE job_id = ?", (job_id,))
        cursor.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
        return cursor.rowcount > 0


# Model metrics operations

def save_model_metrics(job_id: str, metrics: List[Dict[str, Any]]):
    """Save model metrics for preview"""
    with get_db() as conn:
        cursor = conn.cursor()
        # Clear existing metrics
        cursor.execute("DELETE FROM model_metrics WHERE job_id = ?", (job_id,))
        # Insert new metrics
        for metric in metrics:
            cursor.execute("""
                INSERT INTO model_metrics (job_id, metric_name, metric_value, metric_format)
                VALUES (?, ?, ?, ?)
            """, (
                job_id,
                metric.get('name'),
                metric.get('value'),
                metric.get('format', 'number')
            ))


def get_model_metrics(job_id: str) -> List[Dict[str, Any]]:
    """Get model metrics for preview"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT metric_name, metric_value, metric_format 
            FROM model_metrics WHERE job_id = ?
        """, (job_id,))
        return [dict(row) for row in cursor.fetchall()]


# Preferences operations

def get_preference(key: str, default: Any = None) -> Any:
    """Get a preference value"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM preferences WHERE key = ?", (key,))
        row = cursor.fetchone()
        if row:
            try:
                return json.loads(row['value'])
            except json.JSONDecodeError:
                return row['value']
    return default


def set_preference(key: str, value: Any):
    """Set a preference value"""
    with get_db() as conn:
        cursor = conn.cursor()
        json_value = json.dumps(value) if not isinstance(value, str) else value
        cursor.execute("""
            INSERT OR REPLACE INTO preferences (key, value) VALUES (?, ?)
        """, (key, json_value))


# Saved Projects operations

def init_projects_table():
    """Initialize saved projects table"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS saved_projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                description TEXT,
                project_type TEXT DEFAULT 'general',
                configuration TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()


def save_project(
    name: str,
    configuration: Dict[str, Any],
    project_type: str = "general",
    description: str = None
) -> Dict[str, Any]:
    """Save a project configuration"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO saved_projects (name, description, project_type, configuration)
            VALUES (?, ?, ?, ?)
        """, (name, description, project_type, json.dumps(configuration)))
        project_id = cursor.lastrowid
    return get_project(project_id)


def get_project(project_id: int) -> Optional[Dict[str, Any]]:
    """Get a saved project by ID"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM saved_projects WHERE id = ?", (project_id,))
        row = cursor.fetchone()
        if row:
            result = dict(row)
            try:
                result['configuration'] = json.loads(result['configuration'])
            except (json.JSONDecodeError, TypeError):
                pass
            return result
    return None


def get_all_projects() -> List[Dict[str, Any]]:
    """Get all saved projects"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, name, description, project_type, created_at, updated_at 
            FROM saved_projects ORDER BY updated_at DESC
        """)
        return [dict(row) for row in cursor.fetchall()]


def update_project(project_id: int, name: str = None, configuration: Dict = None, description: str = None) -> Optional[Dict[str, Any]]:
    """Update a saved project"""
    updates = []
    values = []
    
    if name:
        updates.append("name = ?")
        values.append(name)
    if configuration:
        updates.append("configuration = ?")
        values.append(json.dumps(configuration))
    if description is not None:
        updates.append("description = ?")
        values.append(description)
    
    if not updates:
        return get_project(project_id)
    
    updates.append("updated_at = CURRENT_TIMESTAMP")
    values.append(project_id)
    
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(f"""
            UPDATE saved_projects SET {", ".join(updates)} WHERE id = ?
        """, values)
    
    return get_project(project_id)


def delete_project(project_id: int) -> bool:
    """Delete a saved project"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM saved_projects WHERE id = ?", (project_id,))
        return cursor.rowcount > 0


# Initialize database on import
init_db()
init_projects_table()
