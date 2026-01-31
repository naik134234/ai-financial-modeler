"""
Job Manager - Persistent job storage using database
Provides a dict-like interface for job management with SQLite persistence
"""

from typing import Dict, Any, Optional, List
from datetime import datetime
import database as db


class JobManager:
    """
    Dict-like interface for job management with database persistence.
    Can be used as a drop-in replacement for the in-memory jobs dict.
    """
    
    def __init__(self):
        # In-memory cache for active jobs (improves performance)
        self._cache: Dict[str, Dict[str, Any]] = {}
    
    def __contains__(self, job_id: str) -> bool:
        """Check if job exists"""
        if job_id in self._cache:
            return True
        job = db.get_job(job_id)
        return job is not None
    
    def __getitem__(self, job_id: str) -> Dict[str, Any]:
        """Get job by ID"""
        # Check cache first
        if job_id in self._cache:
            return self._cache[job_id]
        
        # Load from database
        job = db.get_job(job_id)
        if job is None:
            raise KeyError(f"Job {job_id} not found")
        
        # Convert database row to dict format expected by main.py
        job_dict = self._db_to_dict(job)
        self._cache[job_id] = job_dict
        return job_dict
    
    def __setitem__(self, job_id: str, job_data: Dict[str, Any]) -> None:
        """Create or update a job"""
        # Check if job exists
        existing = db.get_job(job_id)
        
        if existing is None:
            # Create new job
            db.create_job(
                job_id=job_id,
                company_name=job_data.get('company_name', 'Unknown'),
                symbol=job_data.get('request', {}).get('symbol'),
                industry=job_data.get('industry'),
                forecast_years=job_data.get('request', {}).get('forecast_years', 5),
                source=job_data.get('request', {}).get('source', 'stock'),
                request_data=job_data.get('request')
            )
        
        # Update job in database
        update_fields = {}
        
        if 'status' in job_data:
            update_fields['status'] = job_data['status']
        if 'progress' in job_data:
            update_fields['progress'] = job_data['progress']
        if 'message' in job_data:
            update_fields['message'] = job_data['message']
        if 'company_name' in job_data:
            update_fields['company_name'] = job_data['company_name']
        if 'industry' in job_data:
            update_fields['industry'] = job_data['industry']
        if 'file_path' in job_data:
            update_fields['file_path'] = job_data['file_path']
        if 'model_type' in job_data:
            update_fields['model_type'] = job_data['model_type']
        
        if job_data.get('status') == 'completed':
            update_fields['completed_at'] = datetime.now().isoformat()
        
        # Store extra data as result_data JSON
        extra_data = {}
        for key in ['download_url', 'filename', 'validation', 'lbo_summary', 'ma_summary']:
            if key in job_data:
                extra_data[key] = job_data[key]
        if extra_data:
            update_fields['result_data'] = extra_data
        
        if update_fields:
            db.update_job(job_id, **update_fields)
        
        # Update cache
        self._cache[job_id] = job_data
    
    def get(self, job_id: str, default: Any = None) -> Optional[Dict[str, Any]]:
        """Get job with default value if not found"""
        try:
            return self[job_id]
        except KeyError:
            return default
    
    def _db_to_dict(self, db_row: Dict[str, Any]) -> Dict[str, Any]:
        """Convert database row to dict format expected by main.py"""
        import json
        
        result = {
            'job_id': db_row['id'],
            'status': db_row['status'],
            'progress': db_row['progress'] or 0,
            'message': db_row['message'] or '',
            'company_name': db_row['company_name'],
            'industry': db_row['industry'],
            'created_at': db_row['created_at'],
            'file_path': db_row['file_path'],
            'model_type': db_row.get('model_type', 'general'),
        }
        
        # Parse request_data JSON
        if db_row.get('request_data'):
            try:
                result['request'] = json.loads(db_row['request_data'])
            except (json.JSONDecodeError, TypeError):
                result['request'] = {}
        
        # Parse result_data JSON  
        if db_row.get('result_data'):
            try:
                extra = json.loads(db_row['result_data'])
                result.update(extra)
            except (json.JSONDecodeError, TypeError):
                pass
        
        # Add download_url if file exists and completed
        if result['status'] == 'completed' and result.get('file_path'):
            result['download_url'] = f"/api/download/{db_row['id']}"
            if result.get('file_path'):
                import os
                result['filename'] = os.path.basename(result['file_path'])
        
        return result
    
    def get_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Get job history"""
        jobs = db.get_job_history(limit=limit)
        return [self._db_to_dict(job) for job in jobs]
    
    def clear_cache(self, job_id: str = None):
        """Clear cache for a specific job or all jobs"""
        if job_id:
            self._cache.pop(job_id, None)
        else:
            self._cache.clear()


# Global job manager instance
jobs = JobManager()
