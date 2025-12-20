from fastapi import FastAPI, HTTPException, BackgroundTasks
from pydantic import BaseModel
import os
import hashlib
import re
from dotenv import load_dotenv
from cachetools import TTLCache
from extractor import CopilotResponseProcessor, GeminiCompanyExtractor, get_graph_token, download_file_from_sharepoint

# Load environment variables
load_dotenv()

app = FastAPI(title="Competitor Analysis API")

# Initialize in-memory cache with 48-hour TTL (172800 seconds)
# Max 100 entries to prevent memory issues
analysis_cache = TTLCache(maxsize=100, ttl=172800)

# Configuration
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
DRIVE_ID = os.getenv("DRIVE_ID")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_API_KEY_BACKUP = os.getenv("GEMINI_API_KEY_BACKUP")

class AnalysisRequest(BaseModel):
    copilot_response: str
    target_company: str

class AnalysisResponse(BaseModel):
    target_company: str
    data_type: str  # 'public_comps' or 'ma_comps'
    
    # For public comps (with verification)
    verified_competitors: list = []
    to_crosscheck: list = []
    verified_count: int = 0
    crosscheck_count: int = 0
    reasoning: str = ""
    
    # For M&A/transaction comps (no verification)
    ma_transactions: list = []  # List of transaction objects
    transaction_count: int = 0
    
    # Common fields
    files_processed: int
    total_files_found: int
    failed_files: list
    
    cached: bool = False  # Indicates if result came from cache

def get_cache_key(target_company: str, file_paths: list) -> str:
    """
    Generate a unique cache key based on target company and file paths.
    Normalizes company names to improve cache hit rate for variations like:
    - Case differences: "Acme Corp" vs "ACME CORP"
    - Punctuation: "J.P. Morgan" vs "JP Morgan"
    - Suffixes: "Acme Corp" vs "Acme Corporation"
    - Spacing: "Acme Corp" vs "AcmeCorp"
    """
    # Sort file paths to ensure consistent hashing
    paths_str = "".join(sorted(file_paths))
    paths_hash = hashlib.md5(paths_str.encode()).hexdigest()[:8]
    
    # Enhanced company name normalization
    company_key = target_company.lower()
    
    # Replace all non-alphanumeric characters with underscores
    company_key = re.sub(r'[^a-z0-9]', '_', company_key)
    
    # Collapse multiple consecutive underscores into one
    company_key = re.sub(r'_+', '_', company_key)
    
    # Remove leading/trailing underscores
    company_key = company_key.strip('_')
    
    # Remove common business suffixes to improve matching
    # Order matters: check longer suffixes first
    suffixes = [
        '_corporation', '_incorporated', '_limited', 
        '_corp', '_inc', '_llc', '_ltd', '_plc',
        '_the', '_group', '_company', '_co'
    ]
    for suffix in suffixes:
        if company_key.endswith(suffix):
            company_key = company_key[:-len(suffix)].rstrip('_')
            break  # Only remove one suffix
    
    return f"{company_key}_{paths_hash}"

@app.get("/")
def read_root():
    return {"status": "online", "service": "Competitor Analysis API"}

@app.get("/cache/stats")
def get_cache_stats():
    """Get cache statistics"""
    return {
        "cache_size": len(analysis_cache),
        "max_size": analysis_cache.maxsize,
        "ttl_hours": analysis_cache.ttl / 3600,
        "entries": list(analysis_cache.keys())
    }

@app.delete("/cache")
def clear_cache():
    """Clear the entire cache"""
    analysis_cache.clear()
    return {"status": "Cache cleared", "cache_size": 0}

@app.post("/analyze", response_model=AnalysisResponse)
def analyze_competitors(request: AnalysisRequest):
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET, DRIVE_ID, GEMINI_API_KEY]):
        raise HTTPException(status_code=500, detail="Server configuration error: Missing environment variables.")

    try:
        # 1. Initialize Processor
        processor = CopilotResponseProcessor(
            copilot_response=request.copilot_response,
            target_company=request.target_company,
            api_key=GEMINI_API_KEY
        )

        # 2. Extract File Paths
        file_paths, relative_paths = processor.extract_file_paths()
        
        if not file_paths:
            return AnalysisResponse(
                target_company=request.target_company,
                data_type='public_comps',
                verified_competitors=[],
                to_crosscheck=[],
                verified_count=0,
                crosscheck_count=0,
                reasoning="No file paths found in Copilot response.",
                files_processed=0,
                total_files_found=0,
                failed_files=[],
                cached=False
            )

        # 3. Check Cache
        cache_key = get_cache_key(request.target_company, file_paths)
        if cache_key in analysis_cache:
            cached_result = analysis_cache[cache_key]
            
            # Check if the cached result was a "silent failure"
            # 1. Total file failure: All files failed (partial failure is okay)
            all_files_failed = (len(cached_result.failed_files) == cached_result.total_files_found) and (cached_result.total_files_found > 0)
            
            # 2. significant data missing based on type
            data_missing = False
            dtype = cached_result.data_type
            
            if dtype == 'public_comps':
                # invalid if no competitors found at all
                data_missing = (cached_result.verified_count == 0 and cached_result.crosscheck_count == 0)
            elif dtype == 'ma_comps':
                # invalid if no transactions found
                data_missing = (cached_result.transaction_count == 0)
            elif dtype == 'both':
                # invalid if EVERYTHING is missing
                data_missing = (cached_result.verified_count == 0 and cached_result.crosscheck_count == 0 and cached_result.transaction_count == 0)
            
            is_silent_failure = all_files_failed or data_missing
            
            if is_silent_failure:
                print(f"[WARN] Ignoring cached result for {request.target_company} (key: {cache_key})")
                print(f"   Reason: All files failed? {all_files_failed}, Data missing? {data_missing} (Type: {dtype})")
                # checking logic proceeds to process...
            else:
                print(f"[HIT] Cache hit for {request.target_company} (key: {cache_key})")
                cached_result.cached = True
                return cached_result
        
        print(f"[MISS] Cache miss for {request.target_company} (key: {cache_key}). Processing...")

        # 3. Authenticate with SharePoint
        try:
            access_token = get_graph_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"SharePoint Authentication failed: {str(e)}")

        # 4. Process Files
        all_extraction_results = []
        failed_files = []
        processed_count = 0
        
        # Usage Accumulators
        total_input_tokens = 0
        total_output_tokens = 0
        total_input_chars = 0
        per_file_usage = []
        
        for rel_path, full_path in zip(relative_paths, file_paths):
            try:
                # Download file
                file_stream = download_file_from_sharepoint(access_token, DRIVE_ID, rel_path)
                file_size = file_stream.getbuffer().nbytes
                
                # Extract companies or transactions
                extractor = GeminiCompanyExtractor(
                    source=file_stream, 
                    api_key=GEMINI_API_KEY,
                    target_company=request.target_company,
                    backup_api_key=GEMINI_API_KEY_BACKUP
                )
                
                results = extractor.extract_with_gemini()
                
                if results:
                    all_extraction_results.append(results)
                    processed_count += 1
                    
                    # Accumulate Usage - REMOVED
                    pass
                    
                    # Update file logging info
                    ma_count = len(results.get('ma_transactions', []))
                    pub_verified = len(results.get('public_comps', {}).get('verified', []))
                    pub_check = len(results.get('public_comps', {}).get('to_crosscheck', []))
                    processor.file_wise_companies[full_path] = f"Extracted: {ma_count} M&A, {pub_verified} Verified, {pub_check} Check"
                    
                else:
                    failed_files.append({"path": rel_path, "error": "No results returned"})
            except Exception as e:
                print(f"[ERROR] ERROR processing file {rel_path}: {str(e)}")
                failed_files.append({"path": rel_path, "error": str(e)})
                continue
        
        # 5. Aggregate and Build Response
        aggregated = processor.aggregate_unified_results(all_extraction_results)
        
        # Determine overall data type based on existence of data
        has_ma = aggregated['ma_count'] > 0
        has_public = aggregated['verified_count'] > 0 or aggregated['crosscheck_count'] > 0

        if processed_count == 0 and len(failed_files) > 0:
            final_type = 'error'
        elif has_ma and has_public:
            final_type = 'both'
        elif has_ma:
            final_type = 'ma_comps'
        else:
            final_type = 'public_comps'

        # Helper to sort and limit M&A transactions (consistent with previous logic)
        def process_ma_transactions(transactions):
            # Sort by number of non-null metrics
            def count_metrics(t):
                metrics = ['revenue', 'valuation', 'ev_revenue', 'ev_ebitda']
                count = 0
                for m in metrics:
                    val = t.get(m)
                    if val and str(val).lower() != 'null':
                        count += 1
                return count
            
            transactions.sort(key=count_metrics, reverse=True)
            return transactions[:20]

        final_ma_txs = process_ma_transactions(aggregated['ma_transactions'])

        result = AnalysisResponse(
            target_company=request.target_company,
            data_type=final_type,
            
            # M&A Data
            ma_transactions=final_ma_txs,
            transaction_count=len(final_ma_txs),
            
            # Public Comps Data
            verified_competitors=aggregated['verified_competitors'],
            to_crosscheck=aggregated['to_crosscheck'],
            verified_count=aggregated['verified_count'],
            crosscheck_count=aggregated['crosscheck_count'],
            
            # Metadata
            reasoning="Unified Extraction with Unified Reasoning", # Simplified reasoning
            files_processed=processed_count,
            total_files_found=len(file_paths),
            failed_files=failed_files,
            cached=False
        )
        
        # 6. Store in cache (ONLY if not error)
        if final_type != 'error':
            analysis_cache[cache_key] = result
            print(f"[SAVED] Cached result for {request.target_company} (key: {cache_key})")
        else:
            print(f"[SKIP-CACHE] Result type is 'error', not caching.")
        
        return result

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Analysis failed: {str(e)}")
