import requests
import json

# API endpoint
API_URL = "http://127.0.0.1:8000/analyze"

# Load test request
with open("test_request.json", "r") as f:
    request_data = json.load(f)

print("🚀 Sending request to API...")
print(f"Target Company: {request_data['target_company']}")
print(f"Copilot Response Length: {len(request_data['copilot_response'])} characters")
print("NOTE: Check that the server is running and monitor the server terminal for real-time logs (SharePoint/Gemini status).")
print("\n" + "="*60 + "\n")

try:
    # Send POST request (10 minute timeout)
    response = requests.post(API_URL, json=request_data, timeout=600)
    
    print(f"✅ Response Status: {response.status_code}\n")
    
    if response.status_code == 200:
        result = response.json()
        
        print("📊 ANALYSIS RESULTS:")
        print("="*60)
        print(f"Target Company: {result['target_company']}")
        print(f"Data Type: {result['data_type']}")
        print(f"Files Processed: {result['files_processed']}/{result['total_files_found']}")
        print(f"Cached: {result['cached']}")
        print()
        
        # Display based on data type
        if result['data_type'] in ['public_comps', 'both']:
            print(f"Verified Competitors: {result['verified_count']}")
            if result['verified_competitors']:
                print("Sample Verified Competitors:")
                for comp in result['verified_competitors'][:5]:
                    print(f"  • {comp['name']} (Score: {comp['score']})")
            print()
            
            print(f"To Cross-Check: {result['crosscheck_count']}")
            if result['to_crosscheck']:
                print("Sample Cross-Check Candidates:")
                for comp in result['to_crosscheck'][:5]:
                    print(f"  • {comp['name']} (Score: {comp['score']})")
            print()
        
        if result['data_type'] in ['ma_comps', 'both']:
            print(f"M&A Transactions: {result['transaction_count']}")
            if result['ma_transactions']:
                print("Sample M&A Transactions:")
                for txn in result['ma_transactions'][:3]:
                    acq_type = txn.get('acquisition_type', 'Unknown')
                    valuation = txn.get('valuation', 'N/A')
                    print(f"  • {txn['target']} <- {txn['acquirer']} ({valuation}) [{acq_type}]")
            print()
        
        if result['failed_files']:
            print(f"⚠️ Failed Files: {len(result['failed_files'])}")
            for fail in result['failed_files']:
                print(f"  • {fail['path']}: {fail['error']}")
            print()
            
        if 'files_processed' in result:
            print(f"Files Processed: {result['files_processed']} / {result['total_files_found']}")
            print()
        
        print("="*60)
        
        # Save full response
        with open("test_response.json", "w") as f:
            json.dump(result, f, indent=2)
        print("✅ Full response saved to: test_response.json")
        
    else:
        print(f"❌ Error: {response.status_code}")
        print(response.text)
        
except requests.exceptions.Timeout:
    print("❌ Request timed out after 5 minutes")
except requests.exceptions.ConnectionError:
    print("❌ Could not connect to API. Is the server running?")
except Exception as e:
    print(f"❌ Error: {str(e)}")
