"""
AI Financial Modeling Platform - Main API Server
FastAPI backend for generating institutional-grade Excel financial models
"""

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any
import os
import logging
from datetime import datetime
import uuid

# Import our modules
from data import fetch_stock_data, fetch_screener_data
from data.stock_database import get_all_stocks, get_stocks_by_sector, search_stocks, get_sectors
from agents import classify_company, create_model_structure, validate_financial_model
from excel import generate_financial_model

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="AI Financial Modeling Platform",
    description="Generate institutional-grade Excel financial models with AI",
    version="1.0.0",
)

# CORS middleware for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://127.0.0.1:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Output directory for generated models
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# Pydantic models for API
class ModelRequest(BaseModel):
    """Request to generate a financial model"""
    symbol: str = Field(..., description="Stock symbol (e.g., ADANIPOWER)")
    exchange: str = Field(default="NSE", description="Exchange: NSE or BSE")
    forecast_years: int = Field(default=5, ge=1, le=10, description="Number of forecast years")
    model_types: List[str] = Field(
        default=["three_statement", "dcf"],
        description="Model types to include"
    )

class ModelResponse(BaseModel):
    """Response with model generation status"""
    job_id: str
    status: str
    message: str
    company_name: Optional[str] = None
    industry: Optional[str] = None
    download_url: Optional[str] = None

class CompanyInfo(BaseModel):
    """Company information response"""
    symbol: str
    name: str
    sector: str
    industry: str
    market_cap: Optional[float] = None
    current_price: Optional[float] = None

class ValidationResult(BaseModel):
    """Model validation result"""
    is_valid: bool
    errors: List[Dict[str, Any]]

class RawDataRequest(BaseModel):
    """Request to generate model from raw data input"""
    company_name: str = Field(..., description="Name of the company")
    industry: str = Field(default="general", description="Industry code")
    forecast_years: int = Field(default=5, ge=1, le=10)
    historical_data: Dict[str, Any] = Field(
        default={},
        description="Historical financial data (revenue, ebitda, net_income, etc.)"
    )
    assumptions: Dict[str, float] = Field(
        default={},
        description="Override default assumptions (revenue_growth, ebitda_margin, etc.)"
    )


# In-memory job storage (use Redis/DB in production)
jobs: Dict[str, Dict[str, Any]] = {}


@app.get("/")
async def root():
    """Root endpoint with API info"""
    return {
        "name": "AI Financial Modeling Platform",
        "version": "1.0.0",
        "status": "running",
        "endpoints": {
            "company_info": "/api/company/{symbol}",
            "generate_model": "/api/model/generate",
            "job_status": "/api/job/{job_id}",
            "download_model": "/api/download/{job_id}",
        }
    }


@app.get("/api/company/{symbol}", response_model=CompanyInfo)
async def get_company_info(symbol: str, exchange: str = "NSE"):
    """
    Get company information for a stock symbol
    
    Args:
        symbol: Stock symbol (e.g., ADANIPOWER, RELIANCE)
        exchange: NSE or BSE
    """
    try:
        # Fetch data from Yahoo Finance
        data = fetch_stock_data(symbol, exchange)
        info = data.get('company_info', {})
        metrics = data.get('key_metrics', {})
        
        return CompanyInfo(
            symbol=symbol.upper(),
            name=info.get('name', symbol),
            sector=info.get('sector', 'Unknown'),
            industry=info.get('industry', 'Unknown'),
            market_cap=info.get('market_cap'),
            current_price=metrics.get('current_price'),
        )
    except Exception as e:
        logger.error(f"Error fetching company info for {symbol}: {e}")
        raise HTTPException(status_code=404, detail=f"Company not found: {symbol}")


@app.post("/api/model/generate", response_model=ModelResponse)
async def generate_model(request: ModelRequest, background_tasks: BackgroundTasks):
    """
    Generate a financial model for a company
    
    This starts an async job and returns a job ID for tracking.
    """
    job_id = str(uuid.uuid4())
    
    # Initialize job
    jobs[job_id] = {
        "status": "pending",
        "created_at": datetime.now().isoformat(),
        "request": request.dict(),
        "progress": 0,
        "message": "Job queued",
    }
    
    # Start background task
    background_tasks.add_task(
        _generate_model_task,
        job_id,
        request.symbol,
        request.exchange,
        request.forecast_years,
    )
    
    return ModelResponse(
        job_id=job_id,
        status="pending",
        message="Model generation started. Check job status for progress.",
    )


async def _generate_model_task(
    job_id: str,
    symbol: str,
    exchange: str,
    forecast_years: int,
):
    """Background task to generate the financial model"""
    try:
        jobs[job_id]["status"] = "processing"
        jobs[job_id]["progress"] = 10
        jobs[job_id]["message"] = "Fetching financial data..."
        
        # Step 1: Fetch data from multiple sources
        logger.info(f"Fetching data for {symbol}")
        yahoo_data = fetch_stock_data(symbol, exchange)
        
        jobs[job_id]["progress"] = 25
        jobs[job_id]["message"] = "Scraping additional data from Screener..."
        
        try:
            screener_data = fetch_screener_data(symbol)
        except Exception as e:
            logger.warning(f"Screener scraping failed: {e}")
            screener_data = {}
        
        # Merge data sources
        financial_data = {**yahoo_data}
        if screener_data:
            # Add Screener data to supplement Yahoo data
            if 'annual_results' in screener_data:
                financial_data['screener_annual'] = screener_data['annual_results']
            if 'peers' in screener_data:
                financial_data['peers'] = screener_data['peers']
        
        jobs[job_id]["progress"] = 40
        jobs[job_id]["message"] = "Classifying industry..."
        
        # Step 2: Classify industry
        company_info = financial_data.get('company_info', {})
        industry_info = classify_company(company_info)
        
        jobs[job_id]["company_name"] = company_info.get('name', symbol)
        jobs[job_id]["industry"] = industry_info.get('industry_name', 'Unknown')
        jobs[job_id]["progress"] = 55
        jobs[job_id]["message"] = "Designing model structure..."
        
        # Step 3: Design model structure
        model_structure = create_model_structure(
            company_name=company_info.get('name', symbol),
            industry_info=industry_info,
            historical_data=financial_data,
            forecast_years=forecast_years,
        )
        
        jobs[job_id]["progress"] = 70
        jobs[job_id]["message"] = "Generating Excel model..."
        
        # Step 4: Generate Excel file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{symbol}_{industry_info['model_type']}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        generate_financial_model(
            company_name=company_info.get('name', symbol),
            model_structure=model_structure,
            financial_data=financial_data,
            industry_info=industry_info,
            output_path=output_path,
        )
        
        jobs[job_id]["progress"] = 90
        jobs[job_id]["message"] = "Validating model..."
        
        # Step 5: Validate the model
        validation_data = {
            'income_statement': financial_data.get('income_statement', {}),
            'balance_sheet': financial_data.get('balance_sheet', {}),
            'cash_flow': financial_data.get('cash_flow', {}),
            'assumptions': {},
        }
        is_valid, errors = validate_financial_model(
            validation_data,
            industry_info.get('industry_code', 'general')
        )
        
        # Complete
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["progress"] = 100
        jobs[job_id]["message"] = "Model generated successfully!"
        jobs[job_id]["file_path"] = output_path
        jobs[job_id]["filename"] = filename
        jobs[job_id]["validation"] = {
            "is_valid": is_valid,
            "errors": errors[:5] if errors else [],  # Limit to first 5 errors
        }
        jobs[job_id]["download_url"] = f"/api/download/{job_id}"
        
        logger.info(f"Model generated successfully: {filename}")
        
    except Exception as e:
        logger.error(f"Error generating model for job {job_id}: {e}")
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["message"] = f"Error: {str(e)}"
        jobs[job_id]["progress"] = 0


@app.get("/api/job/{job_id}")
async def get_job_status(job_id: str):
    """Get the status of a model generation job"""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    
    response = {
        "job_id": job_id,
        "status": job["status"],
        "progress": job["progress"],
        "message": job["message"],
        "company_name": job.get("company_name"),
        "industry": job.get("industry"),
    }
    
    if job["status"] == "completed":
        response["download_url"] = job.get("download_url")
        response["validation"] = job.get("validation")
    
    return response


@app.get("/api/download/{job_id}")
async def download_model(job_id: str):
    """Download the generated Excel model"""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    
    if job["status"] != "completed":
        raise HTTPException(status_code=400, detail="Model not ready for download")
    
    file_path = job.get("file_path")
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=file_path,
        filename=job.get("filename", "financial_model.xlsx"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/industries")
async def get_industries():
    """Get list of supported industries"""
    from agents.industry_classifier import INDUSTRY_TEMPLATES
    
    industries = []
    for code, template in INDUSTRY_TEMPLATES.items():
        industries.append({
            "code": code,
            "name": template["name"],
            "model_type": template["model_type"],
            "key_metrics": template["key_metrics"],
        })
    
    return {"industries": industries}


@app.get("/api/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
    }


# ======================= STOCK DATABASE ENDPOINTS =======================

@app.get("/api/stocks")
async def list_all_stocks(sector: Optional[str] = None):
    """Get all available stocks, optionally filtered by sector"""
    if sector:
        stocks = get_stocks_by_sector(sector)
    else:
        stocks = get_all_stocks()
    
    return {
        "count": len(stocks),
        "stocks": stocks,
    }


@app.get("/api/stocks/search/{query}")
async def search_for_stocks(query: str):
    """Search stocks by symbol or name"""
    results = search_stocks(query)
    return {
        "count": len(results),
        "results": results,
    }


@app.get("/api/sectors")
async def list_sectors():
    """Get all available sectors"""
    sectors = get_sectors()
    return {
        "sectors": sectors,
        "count": len(sectors),
    }


# ======================= RAW DATA MODEL GENERATION =======================

@app.post("/api/model/generate-raw")
async def generate_model_from_raw_data(request: RawDataRequest, background_tasks: BackgroundTasks):
    """
    Generate a financial model from raw data input
    
    This allows users to provide their own financial data instead of scraping.
    """
    job_id = str(uuid.uuid4())
    
    # Initialize job
    jobs[job_id] = {
        "status": "pending",
        "created_at": datetime.now().isoformat(),
        "request": request.dict(),
        "progress": 0,
        "message": "Job queued",
    }
    
    # Start background task
    background_tasks.add_task(
        _generate_model_from_raw_task,
        job_id,
        request.company_name,
        request.industry,
        request.forecast_years,
        request.historical_data,
        request.assumptions,
    )
    
    return ModelResponse(
        job_id=job_id,
        status="pending",
        message="Model generation from raw data started.",
    )


async def _generate_model_from_raw_task(
    job_id: str,
    company_name: str,
    industry: str,
    forecast_years: int,
    historical_data: Dict[str, Any],
    assumptions: Dict[str, float],
):
    """Background task to generate model from raw data"""
    try:
        jobs[job_id]["status"] = "processing"
        jobs[job_id]["progress"] = 20
        jobs[job_id]["message"] = "Processing raw data..."
        
        # Build financial data structure
        financial_data = {
            'company_info': {
                'name': company_name,
                'sector': industry,
                'industry': industry,
            },
            'income_statement': historical_data.get('income_statement', {}),
            'balance_sheet': historical_data.get('balance_sheet', {}),
            'cash_flow': historical_data.get('cash_flow', {}),
            'key_metrics': historical_data.get('key_metrics', {}),
        }
        
        # Apply user assumptions
        if assumptions:
            financial_data['user_assumptions'] = assumptions
        
        jobs[job_id]["progress"] = 40
        jobs[job_id]["message"] = "Building industry template..."
        
        # Get industry info
        from agents.industry_classifier import INDUSTRY_TEMPLATES
        industry_template = INDUSTRY_TEMPLATES.get(industry, INDUSTRY_TEMPLATES['general'])
        
        industry_info = {
            'industry_name': industry_template['name'],
            'industry_code': industry,
            'model_type': industry_template['model_type'],
            'key_metrics': industry_template['key_metrics'],
        }
        
        jobs[job_id]["company_name"] = company_name
        jobs[job_id]["industry"] = industry_info['industry_name']
        jobs[job_id]["progress"] = 55
        jobs[job_id]["message"] = "Creating model structure..."
        
        # Create model structure
        model_structure = create_model_structure(
            company_name=company_name,
            industry_info=industry_info,
            historical_data=financial_data,
            forecast_years=forecast_years,
        )
        
        # Override with user assumptions
        if assumptions:
            for key_assumption in model_structure.get('key_assumptions', []):
                assumption_name = key_assumption['name'].lower().replace(' ', '_')
                if assumption_name in assumptions:
                    key_assumption['default_value'] = assumptions[assumption_name]
        
        jobs[job_id]["progress"] = 70
        jobs[job_id]["message"] = "Generating Excel model..."
        
        # Generate Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = "".join(c for c in company_name if c.isalnum() or c in (' ', '-', '_')).strip()[:30]
        filename = f"{safe_name}_{industry}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        generate_financial_model(
            company_name=company_name,
            model_structure=model_structure,
            financial_data=financial_data,
            industry_info=industry_info,
            output_path=output_path,
        )
        
        jobs[job_id]["progress"] = 90
        jobs[job_id]["message"] = "Validating model..."
        
        # Validate
        validation_data = {
            'income_statement': financial_data.get('income_statement', {}),
            'balance_sheet': financial_data.get('balance_sheet', {}),
            'cash_flow': financial_data.get('cash_flow', {}),
            'assumptions': assumptions,
        }
        is_valid, errors = validate_financial_model(validation_data, industry)
        
        # Complete
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["progress"] = 100
        jobs[job_id]["message"] = "Model generated successfully from raw data!"
        jobs[job_id]["file_path"] = output_path
        jobs[job_id]["filename"] = filename
        jobs[job_id]["validation"] = {
            "is_valid": is_valid,
            "errors": errors[:5] if errors else [],
        }
        jobs[job_id]["download_url"] = f"/api/download/{job_id}"
        
        logger.info(f"Model generated from raw data: {filename}")
        
    except Exception as e:
        logger.error(f"Error generating model from raw data for job {job_id}: {e}")
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["message"] = f"Error: {str(e)}"
        jobs[job_id]["progress"] = 0


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
