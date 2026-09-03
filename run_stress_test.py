from src.verifier import TrustpilotVerifier

INPUT_XLSM = "LG_corrected.xlsm"
OUTPUT_XLSM = "LG_qa_test.xlsm"

def main():
    verifier = TrustpilotVerifier(INPUT_XLSM, OUTPUT_XLSM)
    
    print("\n==================================================")
    print("STARTING 100-ROW CONCURRENT STRESS TEST")
    print("==================================================")
    
    # We pass limit=100 to only run 100 rows
    tally = verifier.concurrent_verify(limit=100, checkpoint_file="checkpoint_qa.json")
    
    print("\n==================================================")
    print("STRESS TEST COMPLETE")
    print("==================================================")
    
    print(f"\nTOTAL ROWS PROCESSED: {tally['Total_Rows']}\n")
    print(f"VERIFIED_MATCH: {tally.get('VERIFIED_MATCH', 0)}")
    print(f"VERIFIED_AND_CORRECTED: {tally.get('VERIFIED_AND_CORRECTED', 0)}")
    print(f"MULTIPLE_POSSIBLE_MATCHES: {tally.get('MULTIPLE_POSSIBLE_MATCHES', 0)}")
    print(f"REVIEW_NOT_FOUND: {tally.get('REVIEW_NOT_FOUND', 0)}")
    print(f"ACCESS_LIMITED: {tally.get('ACCESS_LIMITED', 0)}")
    print(f"SOURCE_DATA_INSUFFICIENT: {tally.get('SOURCE_DATA_INSUFFICIENT', 0)}")
    print(f"RATE_LIMITED: {tally.get('RATE_LIMITED', 0)}\n")
    
    print("FIELD DISCREPANCIES CORRECTED")
    print(f"Raw_text: {tally['Corrections']['Raw_text']}")
    print(f"Rating: {tally['Corrections']['Rating']}")
    print(f"Review_date: {tally['Corrections']['Review_date']}")
    print(f"Reply_date: {tally['Corrections']['Reply_date']}")
    print(f"Support_reply: {tally['Corrections']['Support_reply']}\n")
    
    print(f"TOTAL RETRIES/ERRORS: {tally.get('Total_Retries_Errors', 0)}")
    
    latencies = tally.get('Latencies', [])
    if latencies:
        avg_lat = sum(latencies) / len(latencies)
        sorted_lats = sorted(latencies)
        p95 = sorted_lats[int(len(sorted_lats) * 0.95)]
        print(f"AVERAGE LATENCY: {avg_lat:.2f}s")
        print(f"p95 LATENCY: {p95:.2f}s")
        
        # Estimate total time for 2198 rows
        # Assuming average concurrency of 15 (just a rough guess based on stress test)
        # We can calculate based on effective req/s
        total_time_stress = max(latencies) if len(latencies) < 100 else sum(latencies) / 10  # Very rough
    
    print(f"\nOutput safely written to: {OUTPUT_XLSM}")

if __name__ == "__main__":
    main()
