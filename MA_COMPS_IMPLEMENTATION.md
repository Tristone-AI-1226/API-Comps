# M&A/Transaction Comps Implementation

## Overview

The system now supports **three types of comps analysis**:

1. **Public Comps** (Trading/Equity/Public) - Extracts company names and verifies competitors
2. **M&A/Transaction Comps** (M&A/Transaction/Precedent/Deal) - Extracts transaction pairs with financial metrics and classifies acquisition type
3. **Both** - When a workbook contains both types, **both flows execute simultaneously**


---

## Key Features
Clear Cache - Invoke-RestMethod -Method DELETE -Uri "http://127.0.0.1:8000/cache"
### 1. **Automatic Sheet Detection & Prioritization**

The system automatically detects the type of comps based on sheet names and selects the best sheets for analysis:

#### Selection Logic
- **Hidden Sheets**: Automatically skipped.
- **Prioritization**:
  - Prefers plain names (e.g., "Public Comps") or sector-specific names (e.g., "Public Comps_Sector").
  - Penalizes names with "pitchbook", "capiq", "factset" suffixes.
  - **Limit**: Processes a maximum of **2 sheets** per type (M&A and Public) to optimize token usage.

#### M&A/Transaction Comps
Matches sheet names containing M&A/transaction keywords:
- `M&A comps`, `MA comps`, `M&A Comps`
- `Transaction comps`, `Precedent comps`, `Deal comps`
- `comps M&A`, `comps transaction` (reverse order)
- Any variation with these keywords

**Pattern**: `(m&a|ma|transaction|precedent|deal).*comps|comps.*(m&a|ma|transaction|precedent|deal)`

**Does NOT match**: Just `"comps"` (that's public comps)

#### Public Comps
Matches sheet names for public/trading comps:
- `Equity comps`, `Trading comps`, `Public comps`
- **`comps`** (standalone - treated as public comps)
- Any variation with equity/trading/public keywords

**Pattern**: `(equity|trading|public).*comps|^comps$`

**Key Rule**: A sheet named just **"comps"** is treated as **public comps** ✅


#### **NEW: Both Types in Same Workbook** ✨
If a workbook contains **both** M&A and public comps sheets:
- **Both sheets are processed** (up to 2 of each)
- M&A transactions are extracted (no verification)
- Public comps are extracted and verified
- Results include both transaction data and verified competitors
- Response `data_type` is set to `"both"`


---

### 2. **M&A Data Extraction & Classification**

For M&A/Transaction sheets, the system extracts:

| Field | Description | Example Column Names |
|-------|-------------|---------------------|
| **Target** | Target company name | Target, Company, Target Name |
| **Acquirer** | Acquiring company | Acquirer, Buyer, Purchaser, Bidder |
| **Type** | Transaction type | Type, Deal Type, Description |
| **Acquisition Type** | **NEW**: Classification | **Strategic** (Same industry) or **Financial** (PE/Fund) |
| **Revenue** | Revenue figures | Revenue, Sales, Turnover, LTM Revenue |
| **Valuation** | Enterprise/Deal value | EV, Enterprise Value, Deal Value, Transaction Value |
| **EV/Revenue** | Revenue multiple | EV/Revenue, EV/Sales, Price/Sales |
| **EV/EBITDA** | EBITDA multiple | EV/EBITDA, Price/EBITDA |

**Flexible Column Matching**: The Gemini AI model intelligently identifies columns even if they don't match exactly.

---

### 3. **Response Limiting & Sorting**

To ensure concise and relevant responses:

- **Public Comps**:
  - **Verified Competitors**: Limited to top **20** (sorted by confidence score).
  - **To Cross-Check**: Limited to top **20** (sorted by confidence score).

- **M&A Transactions**:
  - **Sorting**: Transactions are sorted by the availability of financial metrics (Revenue, Valuation, Multiples).
  - **Limit**: Top **20** transactions with the most complete data are returned.

---

## API Response Structure

### Response Model

```json
{
  "target_company": "string",
  "data_type": "public_comps" | "ma_comps",
  
  // For public_comps
  "verified_competitors": [...],
  "to_crosscheck": [...],
  "verified_count": 0,
  "crosscheck_count": 0,
  "reasoning": "...",
  
  // For ma_comps
  "ma_transactions": [
    {
      "target": "Company A",
      "acquirer": "Company B",
      "type": "Acquisition",
      "acquisition_type": "Strategic",
      "revenue": "$500M",
      "valuation": "$2.5B",
      "ev_revenue": "5.0x",
      "ev_ebitda": "12.5x"
    }
  ],
  "transaction_count": 0,
  
  // Common fields
  "files_processed": 1,
  "total_files_found": 1,
  "failed_files": [],
  "cached": false
}
```

---

## Example Responses

### Public Comps Response

```json
{
  "target_company": "Microsoft",
  "data_type": "public_comps",
  "verified_competitors": [
    {"name": "Google", "score": 95, "reason": "Direct competitor in cloud services"},
    {"name": "Amazon", "score": 88, "reason": "Strong competitor in cloud (AWS)"}
  ],
  "to_crosscheck": [
    {"name": "Oracle", "score": 65, "reason": "Moderate overlap in enterprise software"}
  ],
  "verified_count": 2,
  "crosscheck_count": 1,
  "reasoning": "Analysis based on cloud computing and enterprise software markets",
  "files_processed": 1,
  "total_files_found": 1,
  "failed_files": [],
  "cached": false
}
```

### M&A Comps Response

```json
{
  "target_company": "Tech Startup Inc",
  "data_type": "ma_comps",
  "ma_transactions": [
    {
      "target": "CloudTech Solutions",
      "acquirer": "Microsoft Corporation",
      "type": "Strategic Acquisition",
      "acquisition_type": "Strategic",
      "revenue": "$150M",
      "valuation": "$1.2B",
      "ev_revenue": "8.0x",
      "ev_ebitda": "15.2x"
    },
    {
      "target": "DataAnalytics Pro",
      "acquirer": "Blackstone Group",
      "type": "Buyout",
      "acquisition_type": "Financial",
      "revenue": "$200M",
      "valuation": "$1.5B",
      "ev_revenue": "7.5x",
      "ev_ebitda": "14.0x"
    }
  ],
  "transaction_count": 2,
  "verified_competitors": [],
  "to_crosscheck": [],
  "verified_count": 0,
  "crosscheck_count": 0,
  "reasoning": "",
  "files_processed": 1,
  "total_files_found": 1,
  "failed_files": [],
  "cached": false
}
```

### **NEW: Both Types Response** ✨

When a workbook contains both M&A and public comps sheets:

```json
{
  "target_company": "Tech Startup Inc",
  "data_type": "both",
  "ma_transactions": [
    {
      "target": "CloudTech Solutions",
      "acquirer": "Microsoft Corporation",
      "type": "Strategic Acquisition",
      "acquisition_type": "Strategic",
      "revenue": "$150M",
      "valuation": "$1.2B",
      "ev_revenue": "8.0x",
      "ev_ebitda": "15.2x"
    },
    {
      "target": "DataAnalytics Pro",
      "acquirer": "Blackstone Group",
      "type": "Buyout",
      "acquisition_type": "Financial",
      "revenue": "$200M",
      "valuation": "$1.5B",
      "ev_revenue": "7.5x",
      "ev_ebitda": "14.0x"
    }
  ],
  "transaction_count": 2,
  "verified_competitors": [
    {"name": "Salesforce", "score": 92, "reason": "Direct competitor in SaaS"},
    {"name": "Oracle", "score": 85, "reason": "Strong overlap in enterprise software"}
  ],
  "to_crosscheck": [
    {"name": "SAP", "score": 68, "reason": "Moderate overlap in business applications"}
  ],
  "verified_count": 2,
  "crosscheck_count": 1,
  "reasoning": "Analysis based on SaaS and enterprise software markets",
  "files_processed": 1,
  "total_files_found": 1,
  "failed_files": [],
  "cached": false
}
```

---

## Processing Flow

```
1. Extract file paths from Copilot response
2. Download files from SharePoint
3. For each file:
   ├─ Load Excel workbook
   ├─ Detect ALL sheet types (M&A AND/OR Public)
   │  └─ Filter hidden sheets & prioritize best sheets (max 2 each)
   │
   ├─ IF BOTH M&A and Public sheets found:
   │  ├─ Extract M&A transaction pairs + metrics + classify acquisition type
   │  ├─ Extract public comps company names
   │  ├─ Verify public comps with Gemini
   │  └─ Return BOTH (transactions + verified competitors)
   │
   ├─ IF ONLY M&A/Transaction sheet:
   │  ├─ Extract transaction pairs + metrics + classify acquisition type
   │  └─ Return transactions directly (no verification)
   │
   └─ IF ONLY Public comps sheet:
      ├─ Extract company names
      └─ Verify with Gemini (classify as verified/to-crosscheck)

4. Build response based on data type(s) found
   └─ Limit results to top 20 items per category
5. Cache result (48 hours TTL)
6. Return to client
```

---

## Benefits

✅ **Intelligent Detection**: Automatically identifies M&A vs Public comps  
✅ **Smart Sheet Selection**: Prioritizes best sheets and ignores hidden ones  
✅ **Simultaneous Processing**: Handles both types in the same workbook  
✅ **Rich Transaction Data**: Captures target-acquirer pairs with financial metrics  
✅ **Acquisition Classification**: Distinguishes between Strategic and Financial buyers  
✅ **Response Limiting**: Ensures concise, relevant results (Top 20)  
✅ **Unified API**: Single endpoint handles all data types  
✅ **Cached Results**: 48-hour cache for all response types

---

## Technical Implementation

### Files Modified

1. **`extractor.py`**:
   - Updated `_convert_to_dataframe()` to handle hidden sheets and prioritization
   - Updated M&A extraction prompts to include `acquisition_type` classification

2. **`main.py`**:
   - Updated `AnalysisResponse` model logic
   - Implemented response limiting (max 20) and sorting logic
   - Fixed response building flow

### Regex Patterns

```python
# M&A/Transaction comps
# Requires M&A/transaction/precedent/deal keywords
# Does NOT match just "comps"
ma_pattern = re.compile(
    r'(m&a|ma|transaction|precedent|deal).*comps|comps.*(m&a|ma|transaction|precedent|deal)', 
    re.IGNORECASE
)

# Public comps
# Matches equity/trading/public keywords OR just "comps" (^comps$)
public_pattern = re.compile(
    r'(equity|trading|public).*comps|^comps$', 
    re.IGNORECASE
)
```

**Important**: The `^comps$` in the public pattern ensures that a sheet named just "comps" is treated as public comps.


---

## Usage Notes

1. **Multi-Sheet Processing**: Top 2 matching sheets are processed per type.
2. **Both Types Supported**: If a file contains both M&A and public comps sheets, **both are processed**
3. **Null Values**: Missing metrics are returned as `null` in JSON
4. **Currency Preservation**: Currency symbols and units are preserved (e.g., "$500M", "€1.2B")
5. **Caching**: All response types (public_comps, ma_comps, both) are cached with 48-hour TTL
6. **Verification**: Only public comps go through Gemini verification; M&A transactions are returned directly
7. **Limits**: Max 20 verified competitors, 20 cross-check candidates, and 20 M&A transactions.

---

## Future Enhancements

Potential improvements:
- Support for multiple sheet types in a single file
- Custom metric extraction based on user preferences
- Transaction date extraction
- Deal status tracking (completed, pending, terminated)
- Geographic region extraction
