"""
Cache module for financial data
Provides TTL-based caching using SQLite for persistence
"""

import os
import json
import sqlite3
from datetime import datetime, timedelta
from typing import Any, Optional, Dict
from contextlib import contextmanager
import hashlib
import logging

logger = logging.getLogger(__name__)

# Cache database path
CACHE_DB_PATH = os.path.join(os.path.dirname(__file__), "cache.db")

# Default TTL: 24 hours for stock data
DEFAULT_TTL_HOURS = 24


def get_cache_connection():
    """Get cache database connection"""
    conn = sqlite3.connect(CACHE_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@contextmanager
def get_cache_db():
    """Context manager for cache database connections"""
    conn = get_cache_connection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_cache_db():
    """Initialize cache database tables"""
    with get_cache_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS cache (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP NOT NULL,
                hit_count INTEGER DEFAULT 0
            )
        """)
        
        # Create index for expiry queries
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_cache_expires 
            ON cache(expires_at)
        """)
        conn.commit()


def _generate_cache_key(prefix: str, *args) -> str:
    """Generate a cache key from prefix and arguments"""
    key_parts = [prefix] + [str(arg) for arg in args]
    key_string = ":".join(key_parts)
    return hashlib.md5(key_string.encode()).hexdigest()


def get_cached(key: str) -> Optional[Any]:
    """Get a cached value if it exists and hasn't expired"""
    with get_cache_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT value, expires_at FROM cache 
            WHERE key = ? AND expires_at > datetime('now')
        """, (key,))
        row = cursor.fetchone()
        
        if row:
            # Update hit count
            cursor.execute("""
                UPDATE cache SET hit_count = hit_count + 1 WHERE key = ?
            """, (key,))
            
            try:
                return json.loads(row['value'])
            except json.JSONDecodeError:
                return row['value']
    
    return None


def set_cached(key: str, value: Any, ttl_hours: int = DEFAULT_TTL_HOURS) -> None:
    """Set a cached value with TTL"""
    expires_at = datetime.now() + timedelta(hours=ttl_hours)
    
    with get_cache_db() as conn:
        cursor = conn.cursor()
        
        # Serialize value
        if isinstance(value, (dict, list)):
            json_value = json.dumps(value)
        else:
            json_value = json.dumps(value)
        
        cursor.execute("""
            INSERT OR REPLACE INTO cache (key, value, expires_at, hit_count)
            VALUES (?, ?, ?, 0)
        """, (key, json_value, expires_at.isoformat()))


def delete_cached(key: str) -> bool:
    """Delete a cached value"""
    with get_cache_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM cache WHERE key = ?", (key,))
        return cursor.rowcount > 0


def clear_expired() -> int:
    """Clear all expired cache entries"""
    with get_cache_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM cache WHERE expires_at <= datetime('now')")
        return cursor.rowcount


def clear_all() -> int:
    """Clear entire cache"""
    with get_cache_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM cache")
        return cursor.rowcount


def get_cache_stats() -> Dict[str, Any]:
    """Get cache statistics"""
    with get_cache_db() as conn:
        cursor = conn.cursor()
        
        # Total entries
        cursor.execute("SELECT COUNT(*) as total FROM cache")
        total = cursor.fetchone()['total']
        
        # Expired entries
        cursor.execute("SELECT COUNT(*) as expired FROM cache WHERE expires_at <= datetime('now')")
        expired = cursor.fetchone()['expired']
        
        # Total hits
        cursor.execute("SELECT SUM(hit_count) as hits FROM cache")
        hits = cursor.fetchone()['hits'] or 0
        
        return {
            "total_entries": total,
            "expired_entries": expired,
            "active_entries": total - expired,
            "total_hits": hits,
        }


# Decorator for caching function results
def cached(prefix: str, ttl_hours: int = DEFAULT_TTL_HOURS):
    """Decorator to cache function results"""
    def decorator(func):
        async def wrapper(*args, **kwargs):
            # Generate cache key
            cache_key = _generate_cache_key(prefix, *args, *kwargs.values())
            
            # Try to get from cache
            cached_result = get_cached(cache_key)
            if cached_result is not None:
                logger.info(f"Cache hit for {prefix}")
                return cached_result
            
            # Call function and cache result
            logger.info(f"Cache miss for {prefix}, fetching fresh data")
            result = await func(*args, **kwargs)
            
            if result is not None:
                set_cached(cache_key, result, ttl_hours)
            
            return result
        return wrapper
    return decorator


# Stock data specific caching functions
def cache_stock_data(symbol: str, exchange: str, data: Dict[str, Any], ttl_hours: int = 24) -> None:
    """Cache stock data with symbol-based key"""
    key = _generate_cache_key("stock_data", symbol.upper(), exchange.upper())
    set_cached(key, data, ttl_hours)


def get_cached_stock_data(symbol: str, exchange: str) -> Optional[Dict[str, Any]]:
    """Get cached stock data"""
    key = _generate_cache_key("stock_data", symbol.upper(), exchange.upper())
    return get_cached(key)


def cache_screener_data(symbol: str, data: Dict[str, Any], ttl_hours: int = 24) -> None:
    """Cache screener.in data"""
    key = _generate_cache_key("screener_data", symbol.upper())
    set_cached(key, data, ttl_hours)


def get_cached_screener_data(symbol: str) -> Optional[Dict[str, Any]]:
    """Get cached screener data"""
    key = _generate_cache_key("screener_data", symbol.upper())
    return get_cached(key)


# Initialize cache database on import
init_cache_db()
