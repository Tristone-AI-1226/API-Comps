import pandas as pd
import openpyxl
import io
import os
import requests
import json
import re
from urllib.parse import quote
from msal import ConfidentialClientApplication
import google.generativeai as genai
from typing import List, Dict, Optional, Set

class GeminiCompanyExtractor:
    def __init__(self, source, api_key: str, target_company: str = None, max_sheets: int = 10):
        """
        Initialize the extractor with Gemini integration.
        :param source: file path (str) or BytesIO stream (from SharePoint)
        :param api_key: Google Gemini API Key
        :param target_company: Name of the target company for context-aware extraction
        :param max_sheets: number of sheets to process
        """
        self.source = source
        self.target_company = target_company
        self.max_sheets = max_sheets
        self.results = {}
        
        # Configure Gemini
        genai.configure(api_key=api_key)

    def _convert_to_dataframe(self, workbook):
        """
        Convert Excel workbook to DataFrame, identifying both M&A and public comps sheets.
        Returns: (ma_sheet_data, public_sheet_data, has_ma, has_public)
        """
        ma_sheet_data = {}
        public_sheet_data = {}
        sheet_names = workbook.sheetnames
        
        # Patterns for different comp types
        # M&A pattern: Must have M&A/transaction/precedent/deal keywords
        # Matches: "M&A comps", "Transaction comps", "comps M&A", etc.
        # Does NOT match: just "comps" (that's public)
        ma_pattern = re.compile(r'(m&a|ma|transaction|precedent|deal|private).*comps|comps.*(m&a|ma|transaction|precedent|deal)', re.IGNORECASE)
        
        # Public pattern: equity/trading/public comps OR just "comps"
        # Matches: "Public comps", "Equity comps", "comps", etc.
        public_pattern = re.compile(r'(equity|trading|public).*comps|^comps$', re.IGNORECASE)
        
        ma_sheets = []
        public_sheets = []
        
        # Identify all matching sheets
        for name in sheet_names:
            # Skip hidden sheets
            if workbook[name].sheet_state != 'visible':
                continue

            name_stripped = name.strip()
            
            # Check M&A first (more specific)
            if ma_pattern.search(name_stripped):
                ma_sheets.append(name)
            # Then check public comps
            elif public_pattern.search(name_stripped):
                public_sheets.append(name)

        
        # Helper to score sheet names for prioritization
        def score_sheet_name(name):
            name_lower = name.lower()
            score = 100 # Base score
            
            # Penalize data source suffixes (with variations)
            data_sources = [
                'pitchbook', 'pitch book', 'pitch_book',
                'capiq', 'cap iq', 'cap_iq', 'capitaliq', 'capital iq', 'capital_iq',
                'factset', 'fact set', 'fact_set',
                'crunchbase', 'crunch base', 'crunch_base',
                'preqin', 'pre qin', 'pre_qin'
            ]
            
            for source in data_sources:
                if source in name_lower:
                    score -= 50
                    break  # Only penalize once
            
            # Boost sector/industry keywords
            sector_keywords = [
                'sector', 'industry', 'domain', 'segment', 'category'
            ]
            
            for keyword in sector_keywords:
                if keyword in name_lower:
                    score += 30
                    break  # Only boost once
            
            # Prefer shorter names (usually "Public Comps" vs "Public Comps_Pitchbook")
            score -= len(name)
            
            return score

        # Sort and limit to top 2 sheets
        if ma_sheets:
            ma_sheets_scored = [(s, score_sheet_name(s)) for s in ma_sheets]
            ma_sheets_scored.sort(key=lambda x: x[1], reverse=True)
            ma_sheets = [s for s, _ in ma_sheets_scored[:2]]
        
        if public_sheets:
            public_sheets_scored = [(s, score_sheet_name(s)) for s in public_sheets]
            public_sheets_scored.sort(key=lambda x: x[1], reverse=True)
            public_sheets = [s for s, _ in public_sheets_scored[:2]]
        
        # Process M&A sheets
        for sheet_name in ma_sheets:
            sheet = workbook[sheet_name]
            data = []
            for row in sheet.iter_rows(values_only=True):
                row_list = [str(cell) if cell is not None else "" for cell in row]
                data.append(row_list)
            df = pd.DataFrame(data)
            ma_sheet_data[sheet_name] = df
        
        # Process public comps sheets
        for sheet_name in public_sheets:
            sheet = workbook[sheet_name]
            data = []
            for row in sheet.iter_rows(values_only=True):
                row_list = [str(cell) if cell is not None else "" for cell in row]
                data.append(row_list)
            df = pd.DataFrame(data)
            public_sheet_data[sheet_name] = df
        
        # Fallback: if no specific sheets found, use first 3 visible sheets as public comps
        if not ma_sheets and not public_sheets:
            print(f"⚠️ No 'comps' sheet found. Defaulting to first 3 visible sheets as public comps")
            visible_sheets = [n for n in sheet_names if workbook[n].sheet_state == 'visible']
            for sheet_name in visible_sheets[:3]:
                sheet = workbook[sheet_name]
                data = []
                for row in sheet.iter_rows(values_only=True):
                    row_list = [str(cell) if cell is not None else "" for cell in row]
                    data.append(row_list)
                df = pd.DataFrame(data)
                public_sheet_data[sheet_name] = df
        
        has_ma = len(ma_sheet_data) > 0
        has_public = len(public_sheet_data) > 0
        
        return ma_sheet_data, public_sheet_data, has_ma, has_public

    def _prepare_context_for_gemini(self, sheet_data, max_chars=3200000):
        """
        Prepare Excel data as structured text context for Gemini.
        
        Token Limits (Gemini 2.5 Flash):
        - Input: 1,000,000 tokens (~4M characters)
        - Output: 65,536 tokens
        - Ratio: ~4 characters per token
        
        This function limits to 3.2M chars (~800K tokens = 80% of capacity)
        Uses maximum available capacity while leaving 20% safety margin.
        """
        context = "Below is data from an Excel file containing company information:\n\n"
        
        for sheet_name, df in sheet_data.items():
            sheet_context = f"=== SHEET: {sheet_name} ===\n"
            # With 3.2M char limit, we can handle very large sheets
            # No row limit - let Gemini handle as much data as possible
            sheet_str = df.to_string(index=False, max_rows=None)
            
            # Truncate if still too long
            if len(context) + len(sheet_context) + len(sheet_str) > max_chars:
                remaining_chars = max_chars - len(context) - len(sheet_context)
                if remaining_chars > 0:
                    sheet_str = sheet_str[:remaining_chars] + "\n... (truncated due to size limits)"
                else:
                    print(f"⚠️ Skipping sheet '{sheet_name}' - context limit reached")
                    break  # Skip this sheet if we're already at limit
            
            sheet_context += sheet_str
            sheet_context += "\n\n"
            context += sheet_context
        
        # Final safety check
        if len(context) > max_chars:
            context = context[:max_chars] + "\n... (truncated due to size limits)"
        
        print(f"📊 Context size: {len(context):,} chars (~{len(context)//4:,} tokens, {(len(context)//4)/10000:.1f}% of 1M limit)")
        
        return context

    def extract_ma_transactions_with_gemini(self, model_name='gemini-2.5-flash'):
        """
        Extract M&A transaction data including target-acquirer pairs and metrics.
        Returns structured transaction data instead of just company names.
        """
        try:
            # Load workbook
            if isinstance(self.source, io.BytesIO):
                wb = openpyxl.load_workbook(self.source, data_only=True)
            else:
                file_ext = os.path.splitext(self.source)[1].lower()
                if file_ext == ".csv":
                    df = pd.read_csv(self.source)
                    buf = io.BytesIO()
                    df.to_excel(buf, index=False)
                    buf.seek(0)
                    wb = openpyxl.load_workbook(buf, data_only=True)
                else:
                    wb = openpyxl.load_workbook(self.source, data_only=True)
        except Exception as e:
            print(f"ERROR loading workbook: {e}")
            raise

        # Convert to structured text
        ma_sheet_data, _, has_ma, _ = self._convert_to_dataframe(wb)
        
        if not has_ma:
            wb.close()
            return None
            
        context = self._prepare_context_for_gemini(ma_sheet_data)
        wb.close()

        # Create M&A-specific prompt
        prompt = f"""{context}

TASK: Extract M&A transaction data from the spreadsheet above. DO NOT HALLUCINATE.

IMPORTANT: This is M&A/Transaction/Precedent comps data. Extract transaction pairs and metrics.

Instructions:
1. Identify columns containing:
   - Target company names (may be labeled as: Target, Company, Target Name, etc.)
   - Acquirer/Buyer company names (may be labeled as: Acquirer, Buyer, Purchaser, Bidder, etc.)
   - Transaction type/description (may be labeled as: Type, Deal Type, Description, etc.)
   - Financial metrics(Upto 2 decimal points):
     * Revenue (may be labeled as: Revenue, Sales, Turnover, LTM Revenue, etc.)
     * Valuation/Enterprise Value (may be labeled as: EV, Enterprise Value, Deal Value, Transaction Value, etc.)
     * Value/Revenue multiple (may be labeled as: EV/Revenue, EV/Sales, Price/Sales, etc.)
     * Value/EBITDA multiple (may be labeled as: EV/EBITDA, Price/EBITDA, etc.)

2. CLASSIFY ACQUISITION TYPE:
   - Compare the Acquirer and Target company.
   - If the Acquirer is in the SAME industry as the Target company, classify as "Strategic".
   - If the Acquirer is a Private Equity (PE) fund, investment firm, or financial sponsor, classify as "Financial".
   - If you cannot determine with confidence, use "Unknown".
   - Assign this to the field "acquisition_type".

3. Extract ALL transaction records found in the data
4. Ignore summary rows (Total, Average, Median, Mean, etc.)
5. Ignore rows with N/A, TBD, or missing critical data

6. Return results as a JSON object with this structure:
{{
    "transactions": [
        {{
            "target": "Target Company Name",
            "acquirer": "Acquirer Company Name",
            "type": "Transaction type or description",
            "acquisition_type": "Strategic" | "Financial" | "Unknown",
            "revenue": "Revenue value (with units if available)",
            "valuation": "Enterprise/Deal value (with units if available)",
            "ev_revenue": "EV/Revenue multiple",
            "ev_ebitda": "EV/EBITDA multiple"
        }},
        ...
    ],
    "count": <number of transactions>
}}

7. If a metric is not found or not available, use null for that field
8. Preserve currency symbols and units (e.g., "$500M", "€1.2B")

CRITICAL: Provide ONLY valid JSON response, no additional text, no markdown formatting, no explanations."""

        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            full_response = response.text

            # Parse JSON response
            try:
                # Remove markdown code blocks if present
                full_response = full_response.replace('```json', '').replace('```', '').strip()

                # Extract JSON from response
                json_start = full_response.find('{')
                json_end = full_response.rfind('}') + 1

                if json_start == -1 or json_end == 0:
                    return None

                json_str = full_response[json_start:json_end]
                result = json.loads(json_str)

                if "transactions" not in result:
                    return None

                transactions = result.get("transactions", [])

                self.results = {
                    "transactions": transactions,
                    "total_transactions": len(transactions),
                    "type": "ma_comps"
                }
                return self.results

            except json.JSONDecodeError as e:
                print(f"JSON Decode Error: {e}")
                return None

        except Exception as e:
            print(f"Exception during Gemini API call: {e}")
            return None


    def extract_with_gemini(self, model_name='gemini-2.5-flash'):
        """
        Extract data from Excel using Gemini.
        Handles both M&A transactions and public comps, processing both if present.
        """
        try:
            # Load workbook
            if isinstance(self.source, io.BytesIO):
                wb = openpyxl.load_workbook(self.source, data_only=True)
            else:
                file_ext = os.path.splitext(self.source)[1].lower()
                if file_ext == ".csv":
                    df = pd.read_csv(self.source)
                    buf = io.BytesIO()
                    df.to_excel(buf, index=False)
                    buf.seek(0)
                    wb = openpyxl.load_workbook(buf, data_only=True)
                else:
                    wb = openpyxl.load_workbook(self.source, data_only=True)
        except Exception as e:
            print(f"ERROR loading workbook: {e}")
            raise

        # Detect both M&A and public comps sheets
        ma_sheet_data, public_sheet_data, has_ma, has_public = self._convert_to_dataframe(wb)
        
        results = {}
        
        # Process M&A sheets if present
        if has_ma:
            # Calculate total estimated tokens (approx 4 chars per token)
            total_chars = sum(len(df.to_string()) for df in ma_sheet_data.values())
            estimated_tokens = total_chars // 4
            
            # Gemini Free Tier Limit: 250,000 tokens per minute
            # We'll use a safe batch size of ~200,000 tokens
            BATCH_TOKEN_LIMIT = 200000
            
            if estimated_tokens > BATCH_TOKEN_LIMIT:
                # Split sheets into batches
                batches = []
                current_batch = {}
                current_batch_tokens = 0
                
                for sheet_name, df in ma_sheet_data.items():
                    sheet_tokens = len(df.to_string()) // 4
                    
                    if current_batch_tokens + sheet_tokens > BATCH_TOKEN_LIMIT and current_batch:
                        batches.append(current_batch)
                        current_batch = {}
                        current_batch_tokens = 0
                    
                    current_batch[sheet_name] = df
                    current_batch_tokens += sheet_tokens
                
                if current_batch:
                    batches.append(current_batch)
                
                # Process each batch
                all_transactions = []
                for i, batch in enumerate(batches):
                    batch_context = self._prepare_context_for_gemini(batch)
                    batch_results = self._extract_ma_data(batch_context, model_name)
                    
                    if batch_results and batch_results.get('transactions'):
                        all_transactions.extend(batch_results['transactions'])
                    
                    # Wait between batches to respect rate limit (if not the last one)
                    if i < len(batches) - 1:
                        import time
                        time.sleep(10)
                
                if all_transactions:
                    results['ma_transactions'] = all_transactions
                    results['has_ma'] = True
                    
            else:
                # Process all at once if within limits
                ma_context = self._prepare_context_for_gemini(ma_sheet_data)
                ma_results = self._extract_ma_data(ma_context, model_name)
                if ma_results:
                    results['ma_transactions'] = ma_results.get('transactions', [])
                    results['has_ma'] = True
        
        # Process public comps sheets if present
        if has_public:
            # Calculate total estimated tokens
            total_chars = sum(len(df.to_string()) for df in public_sheet_data.values())
            estimated_tokens = total_chars // 4
            
            # Gemini Free Tier Limit: 250,000 tokens per minute
            # Use safe batch size
            BATCH_TOKEN_LIMIT = 200000
            
            if estimated_tokens > BATCH_TOKEN_LIMIT:
                # Split sheets into batches
                batches = []
                current_batch = {}
                current_batch_tokens = 0
                
                for sheet_name, df in public_sheet_data.items():
                    sheet_tokens = len(df.to_string()) // 4
                    
                    if current_batch_tokens + sheet_tokens > BATCH_TOKEN_LIMIT and current_batch:
                        batches.append(current_batch)
                        current_batch = {}
                        current_batch_tokens = 0
                    
                    current_batch[sheet_name] = df
                    current_batch_tokens += sheet_tokens
                
                if current_batch:
                    batches.append(current_batch)
                
                # Process each batch
                all_companies = set()
                for i, batch in enumerate(batches):
                    batch_context = self._prepare_context_for_gemini(batch)
                    batch_results = self._extract_public_comps_data(batch_context, model_name)
                    
                    if batch_results and batch_results.get('companies'):
                        all_companies.update(batch_results['companies'])
                    
                    # Wait between batches
                    if i < len(batches) - 1:
                        import time
                        time.sleep(10)
                
                if all_companies:
                    results['all_companies'] = all_companies
                    results['has_public'] = True
            
            else:
                # Process all at once if within limits
                public_context = self._prepare_context_for_gemini(public_sheet_data)
                public_results = self._extract_public_comps_data(public_context, model_name)
                if public_results:
                    results['all_companies'] = public_results.get('companies', set())
                    results['has_public'] = True
        
        wb.close()
        
        # Return combined results
        if not results:
            return None
            
        # Set type based on what was found
        if has_ma and has_public:
            results['type'] = 'both'
        elif has_ma:
            results['type'] = 'ma_comps'
        else:
            results['type'] = 'public_comps'
        
        self.results = results
        return results

    def _extract_ma_data(self, context, model_name='gemini-2.5-flash'):
        """Helper method to extract M&A transaction data from context."""
        prompt = f"""{context}

TASK: Extract M&A transaction data from the spreadsheet above. DO NOT HALLUCINATE.

IMPORTANT: This is M&A/Transaction/Precedent comps data. Extract transaction pairs and metrics.

Instructions:
1. Identify columns containing:
   - Target company names (may be labeled as: Target, Company, Target Name, etc.)
   - Acquirer/Buyer company names (may be labeled as: Acquirer, Buyer, Purchaser, Bidder, etc.)
   - Transaction type/description (may be labeled as: Type, Deal Type, Description, etc.)
   - Financial metrics(Upto 2 decimal points):
     * Revenue (may be labeled as: Revenue, Sales, Turnover, LTM Revenue, etc.)
     * Valuation/Enterprise Value (may be labeled as: EV, Enterprise Value, Deal Value, Transaction Value, etc.)
     * Value/Revenue multiple (may be labeled as: EV/Revenue, EV/Sales, Price/Sales, etc.)
     * Value/EBITDA multiple (may be labeled as: EV/EBITDA, Price/EBITDA, etc.)

2. CLASSIFY ACQUISITION TYPE:
   - Compare the Acquirer and Target company.
   - If the Acquirer is in the SAME industry as the Target company, classify as "Strategic".
   - If the Acquirer is a Private Equity (PE) fund, investment firm, or financial sponsor, classify as "Financial".
   - If you cannot determine with confidence, use "Unknown".
   - Assign this to the field "acquisition_type".

3. Extract ALL transaction records found in the data
4. Ignore summary rows (Total, Average, Median, Mean, etc.)
5. Ignore rows with N/A, TBD, or missing critical data

6. Return results as a JSON object with this structure:
{{
    "transactions": [
        {{
            "target": "Target Company Name",
            "acquirer": "Acquirer Company Name",
            "type": "Transaction type or description",
            "acquisition_type": "Strategic" | "Financial" | "Unknown",
            "revenue": "Revenue value (with units if available)",
            "valuation": "Enterprise/Deal value (with units if available)",
            "ev_revenue": "EV/Revenue multiple",
            "ev_ebitda": "EV/EBITDA multiple"
        }},
        ...
    ],
    "count": <number of transactions>
}}

7. If a metric is not found or not available, use null for that field
8. Preserve currency symbols and units (e.g., "$500M", "€1.2B")

CRITICAL: Provide ONLY valid JSON response, no additional text, no markdown formatting, no explanations."""

        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            full_response = response.text

            # Parse JSON response
            full_response = full_response.replace('```json', '').replace('```', '').strip()
            json_start = full_response.find('{')
            json_end = full_response.rfind('}') + 1

            if json_start == -1 or json_end == 0:
                return None

            json_str = full_response[json_start:json_end]
            result = json.loads(json_str)

            if "transactions" not in result:
                return None

            return result

        except Exception as e:
            print(f"Exception during M&A extraction: {e}")
            return None

    def _extract_public_comps_data(self, context, model_name='gemini-2.5-flash'):
        """Helper method to extract public comps company names from context."""
        # Create context-aware prompt
        if self.target_company:
            target_context = f"\nTARGET COMPANY CONTEXT: {self.target_company}\n"
            target_context += f"IMPORTANT: Use your knowledge of {self.target_company}'s industry, products, and market to filter the extracted companies.\n"
            target_context += f"Only extract companies that operate in the SAME or CLOSELY RELATED business as {self.target_company}.\n"
            target_context += f"Exclude companies from completely different industries or product categories.\n\n"
        else:
            target_context = ""

        # Create prompt for Gemini
        prompt = f"""{context}

{target_context}TASK: Extract ALL company names from the data above that are potential competitors or comparable companies.

Instructions:
- Look for columns containing company names, targets, acquirers, sellers, or similar identifiers
- Extract only actual company names (exclude headers, totals, averages, summaries)
- Ignore entries like "N/A", "TBD", "Others", "Mean", "Total", "Average", "Median"
- Include ALL companies found in the spreadsheet
- Return the results as a JSON object with the following structure:
{{
    "companies": ["Company 1", "Company 2", ...],
    "count": <number of unique companies>
}}

CRITICAL: Provide ONLY valid JSON response, no additional text, no markdown formatting, no explanations."""

        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            full_response = response.text

            # Parse JSON response
            full_response = full_response.replace('```json', '').replace('```', '').strip()
            json_start = full_response.find('{')
            json_end = full_response.rfind('}') + 1

            if json_start == -1 or json_end == 0:
                return None

            json_str = full_response[json_start:json_end]
            result = json.loads(json_str)

            if "companies" not in result:
                return None

            companies = result.get("companies", [])
            return {
                "companies": set(companies),
                "count": len(set(companies))
            }

        except json.JSONDecodeError as e:
            return None
        except Exception as e:
            return None


class CopilotResponseProcessor:
    def __init__(self, copilot_response: str, target_company: str, api_key: str):
        """
        Initialize processor with copilot response and target company name.
        """
        self.copilot_response = copilot_response
        self.target_company = target_company
        self.api_key = api_key
        self.file_paths = []
        self.relative_paths = []
        self.all_competitors = set()
        self.verified_competitors = []
        self.to_crosscheck = []
        self.file_wise_companies = {}
        
        # Configure Gemini
        genai.configure(api_key=api_key)

    def extract_file_paths(self):
        """Extract unique file paths from copilot response and cut at file extension"""
        # Pattern to match "Full Path: " followed by the path, cut at file extension
        pattern = r'Full Path:\s*(.+?\.(?:xlsx|xls|csv|pptx|pdf))'
        matches = re.findall(pattern, self.copilot_response, re.IGNORECASE)
        unique_paths = list(set(matches))
        self.file_paths = [path.strip() for path in unique_paths]

        # Extract relative paths starting from "Shared Documents/"
        self.relative_paths = []
        for path in self.file_paths:
            if "Shared Documents/" in path:
                relative_path = path.split("Shared Documents/", 1)[1]
                self.relative_paths.append(relative_path)
            else:
                self.relative_paths.append(path)

        return self.file_paths, self.relative_paths

    def classify_competitors_with_gemini(self, all_companies: list, model_name='gemini-2.5-flash'):
        """
        Use Gemini to classify extracted companies into verified competitors and to-crosscheck.
        
        Token Limits (Gemini 2.5 Flash):
        - Input: 1,000,000 tokens (~4M characters)
        - Using 80% capacity = 800K tokens available
        - This allows for ~2000 companies comfortably
        """
        companies_list = sorted(list(all_companies))
        
        # Limit to 2000 companies (80% capacity usage)
        # Each company name ~30 chars avg, 2000 companies ≈ 60K chars (~15K tokens)
        # Plus prompt (~2K tokens) = ~17K tokens total
        # Still leaves room for 783K tokens of Excel data
        if len(companies_list) > 2000:
            print(f"⚠️ Warning: {len(companies_list)} companies found. Limiting to 2000 for classification.")
            companies_list = companies_list[:2000]
        
        print(f"🔍 Classifying {len(companies_list)} companies for {self.target_company}")

        prompt = f"""You are a business analyst expert specializing in competitive analysis.

TARGET COMPANY: {self.target_company}

EXTRACTED COMPANIES CANDIDATES:
{json.dumps(companies_list, indent=2)}

TASK: Classify these candidates based on their competitive relationship with {self.target_company}.

RULES:
1. **STRICTLY** use ONLY the companies provided in the list above. DO NOT add any new companies.
2. Assign a **Confidence Score (0-100)** representing the strength of the competitive overlap.
   - 90-100: Direct competitor (same core products/services, same market).
   - 70-89: Strong competitor (significant overlap).
   - 50-69: Moderate/Indirect competitor or substitute.
   - <50: Low relevance or different industry.

CLASSIFICATION CATEGORIES:
1. **Verified Competitors**: Score >= 70. Direct/Strong competitors.
2. **To Cross-Check**: Score < 70. Indirect, potential, or unclear competitors.

RESPONSE FORMAT:
Return a JSON object with two lists. Each item must be an object containing "name" and "score".
Sort both lists by "score" in DESCENDING order.

{{
    "verified_competitors": [
        {{"name": "Company A", "score": 95, "reason": "..."}},
        {{"name": "Company B", "score": 88, "reason": "..."}}
    ],
    "to_crosscheck": [
        {{"name": "Company C", "score": 45, "reason": "..."}}
    ],
    "reasoning": "Brief analysis of the industry context."
}}"""

        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            text_response = response.text

            # Remove markdown code blocks if present
            text_response = text_response.replace('```json', '').replace('```', '').strip()

            # Extract JSON from response
            json_start = text_response.find('{')
            json_end = text_response.rfind('}') + 1
            json_str = text_response[json_start:json_end]

            result = json.loads(json_str)

            self.verified_competitors = result.get("verified_competitors", [])
            self.to_crosscheck = result.get("to_crosscheck", [])
            
            return result

        except Exception as e:
            print(f"Exception during classification: {e}")
            # Fallback: put all in to_crosscheck
            self.to_crosscheck = companies_list
            return {
                "verified_competitors": [],
                "to_crosscheck": companies_list,
                "verified_count": 0,
                "crosscheck_count": len(companies_list),
                "reasoning": "Error during classification, fallback to cross-check."
            }

def get_graph_token(tenant_id, client_id, client_secret):
    """Get Microsoft Graph API access token."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    try:
        app = ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret)
        token = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in token:
            return token["access_token"]
        else:
            raise Exception("No access token in response")
    except Exception as e:
        raise Exception(f"Error getting token: {e}")

def search_file_by_name(access_token, drive_id, filename):
    """Search for a file by name in SharePoint if direct path fails."""
    search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='{filename}')"
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        resp = requests.get(search_url, headers=headers)
        resp.raise_for_status()
        results = resp.json()
        if "value" in results and len(results["value"]) > 0:
            return results["value"][0]
        return None
    except Exception:
        return None

def download_file_from_sharepoint(access_token, drive_id, relative_path):
    """Download file from SharePoint using Microsoft Graph API with fallback search."""
    # Try direct path first
    encoded_path = quote(relative_path, safe='/')
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/content"
    headers = {"Authorization": f"Bearer {access_token}"}

    try:
        resp = requests.get(url, headers=headers)

        if resp.status_code == 404:
            filename = os.path.basename(relative_path)
            file_item = search_file_by_name(access_token, drive_id, filename)

            if file_item:
                download_url = file_item.get('@microsoft.graph.downloadUrl')
                if download_url:
                    resp = requests.get(download_url)
                    resp.raise_for_status()
                    return io.BytesIO(resp.content)
                else:
                    raise Exception(f"No download URL for file: {filename}")
            else:
                raise Exception(f"File not found: {filename}")

        resp.raise_for_status()
        return io.BytesIO(resp.content)

    except Exception as e:
        raise Exception(f"Download failed: {e}")
