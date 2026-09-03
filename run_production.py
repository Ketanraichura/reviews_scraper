from src.verifier import TrustpilotVerifier

INPUT_XLSM = "LG_corrected.xlsm"
OUTPUT_XLSM = "LG_audited_final.xlsm"

def main():
    verifier = TrustpilotVerifier(INPUT_XLSM, OUTPUT_XLSM)
    
    print("\n==================================================")
    print("STARTING FULL PRODUCTION RUN (2,198 ROWS)")
    print("==================================================")
    
    # We pass limit=None to run all rows, using a production checkpoint
    tally = verifier.concurrent_verify(limit=None, checkpoint_file="checkpoint_prod.json")
    
    print("\n==================================================")
    print("PRODUCTION RUN COMPLETE")
    print("==================================================")
    
    print(f"\nTOTAL ROWS PROCESSED: {tally.get('Total_Rows', 0)}\n")
    print(f"VERIFIED_MATCH: {tally.get('VERIFIED_MATCH', 0)}")
    print(f"VERIFIED_AND_CORRECTED: {tally.get('VERIFIED_AND_CORRECTED', 0)}")
    print(f"MULTIPLE_POSSIBLE_MATCHES: {tally.get('MULTIPLE_POSSIBLE_MATCHES', 0)}")
    print(f"REVIEW_NOT_FOUND: {tally.get('REVIEW_NOT_FOUND', 0)}")
    print(f"ACCESS_LIMITED: {tally.get('ACCESS_LIMITED', 0)}")
    print(f"SOURCE_DATA_INSUFFICIENT: {tally.get('SOURCE_DATA_INSUFFICIENT', 0)}")
    print(f"RATE_LIMITED: {tally.get('RATE_LIMITED', 0)}\n")
    
    print("FIELD DISCREPANCIES CORRECTED")
    print(f"Raw_text: {tally['Corrections'].get('Raw_text', 0)}")
    print(f"Rating: {tally['Corrections'].get('Rating', 0)}")
    print(f"Review_date: {tally['Corrections'].get('Review_date', 0)}")
    print(f"Reply_date: {tally['Corrections'].get('Reply_date', 0)}")
    print(f"Support_reply: {tally['Corrections'].get('Support_reply', 0)}\n")
    
    print(f"TOTAL RETRIES/ERRORS: {tally.get('Total_Retries_Errors', 0)}")
    print(f"\nOutput safely written to: {OUTPUT_XLSM}")

if __name__ == "__main__":
    main()
