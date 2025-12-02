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
            print(f"✅ Cache hit for {request.target_company} (key: {cache_key})")
            cached_result = analysis_cache[cache_key]
            cached_result.cached = True
            return cached_result
        
        print(f"🔄 Cache miss for {request.target_company} (key: {cache_key}). Processing...")

        # 3. Authenticate with SharePoint
        try:
            access_token = get_graph_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"SharePoint Authentication failed: {str(e)}")

        # 4. Process Files
        all_competitors_set = set()
        all_ma_transactions = []
        failed_files = []
        processed_count = 0
        
        has_ma_data = False
        has_public_data = False

        for rel_path, full_path in zip(relative_paths, file_paths):
            try:
                # Download file
                file_stream = download_file_from_sharepoint(access_token, DRIVE_ID, rel_path)
                
                # Extract companies or transactions
                extractor = GeminiCompanyExtractor(
                    source=file_stream, 
                    api_key=GEMINI_API_KEY,
                    target_company=request.target_company
                )
                
                results = extractor.extract_with_gemini()
                
                if results:
                    result_type = results.get("type")
                    
                    # Handle different result types
                    if result_type == "both":
                        # Both M&A and public comps present
                        has_ma_data = True
                        has_public_data = True
                        
                        # Extract M&A transactions
                        if results.get("ma_transactions"):
                            transactions = results["ma_transactions"]
                            all_ma_transactions.extend(transactions)
                        
                        # Extract public comps
                        if results.get("all_companies"):
                            companies = results["all_companies"]
                            all_competitors_set.update(companies)
                        
                        processor.file_wise_companies[full_path] = f"Both: {len(results.get('ma_transactions', []))} transactions, {len(results.get('all_companies', []))} companies"
                        processed_count += 1
                        
                    elif result_type == "ma_comps":
                        # Only M&A data
                        has_ma_data = True
                        transactions = results.get("ma_transactions", [])
                        all_ma_transactions.extend(transactions)
                        processor.file_wise_companies[full_path] = f"M&A Transactions: {len(transactions)}"
                        processed_count += 1
                        
                    elif result_type == "public_comps" or results.get("all_companies"):
                        # Only public comps data
                        has_public_data = True
                        companies = results.get("all_companies", set())
                        processor.file_wise_companies[full_path] = list(companies)
                        all_competitors_set.update(companies)
                        processed_count += 1
                    else:
                        failed_files.append({"path": rel_path, "error": "No data extracted"})
                else:
                    failed_files.append({"path": rel_path, "error": "No results returned"})
                    
            except Exception as e:
                failed_files.append({"path": rel_path, "error": str(e)})
                continue
        
        # Determine final data type based on aggregated results
        if has_ma_data and has_public_data:
            data_type = 'both'
        elif has_ma_data:
            data_type = 'ma_comps'
        else:
            data_type = 'public_comps'

        # 5. Build Response based on data type
        # 5. Build Response based on data type
        
        # Helper to sort and limit M&A transactions
        def process_ma_transactions(transactions):
            # Sort by number of non-null metrics
            def count_metrics(t):
                metrics = ['revenue', 'valuation', 'ev_revenue', 'ev_ebitda']
                # Check if metric exists and is not None/null string
                count = 0
                for m in metrics:
                    val = t.get(m)
                    if val and str(val).lower() != 'null':
                        count += 1
                return count
            
            # Sort descending by metric count
            transactions.sort(key=count_metrics, reverse=True)
            # Limit to 20
            return transactions[:20]

        if data_type == 'both':
            # Both M&A and public comps - return both with verification for public comps
            classification_result = processor.classify_competitors_with_gemini(list(all_competitors_set)) if all_competitors_set else {}
            
            # Limit lists
            verified = classification_result.get("verified_competitors", [])[:20]
            crosscheck = classification_result.get("to_crosscheck", [])[:20]
            ma_txs = process_ma_transactions(all_ma_transactions)
            
            result = AnalysisResponse(
                target_company=request.target_company,
                data_type='both',
                ma_transactions=ma_txs,
                transaction_count=len(ma_txs),
                verified_competitors=verified,
                to_crosscheck=crosscheck,
                verified_count=len(verified),
                crosscheck_count=len(crosscheck),
                reasoning=classification_result.get("reasoning", ""),
                files_processed=processed_count,
                total_files_found=len(file_paths),
                failed_files=failed_files,
                cached=False
            )
            
        elif data_type == 'ma_comps':
            # Only M&A data
            ma_txs = process_ma_transactions(all_ma_transactions)
            
            result = AnalysisResponse(
                target_company=request.target_company,
                data_type='ma_comps',
                ma_transactions=ma_txs,
                transaction_count=len(ma_txs),
                verified_competitors=[],
                to_crosscheck=[],
                verified_count=0,
                crosscheck_count=0,
                reasoning="",
                files_processed=processed_count,
                total_files_found=len(file_paths),
                failed_files=failed_files,
                cached=False
            )
            
        else:
            # Default to public comps
            classification_result = processor.classify_competitors_with_gemini(list(all_competitors_set)) if all_competitors_set else {}
            
            # Limit lists
            verified = classification_result.get("verified_competitors", [])[:20]
            crosscheck = classification_result.get("to_crosscheck", [])[:20]
            
            result = AnalysisResponse(
                target_company=request.target_company,
                data_type='public_comps',
                verified_competitors=verified,
                to_crosscheck=crosscheck,
                verified_count=len(verified),
                crosscheck_count=len(crosscheck),
                reasoning=classification_result.get("reasoning", ""),
                files_processed=processed_count,
                total_files_found=len(file_paths),
                failed_files=failed_files,
                cached=False
            )
        
        # 6. Store in cache
        analysis_cache[cache_key] = result
        print(f"💾 Cached result for {request.target_company} (key: {cache_key})")
        
        return result

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Analysis failed: {str(e)}")
