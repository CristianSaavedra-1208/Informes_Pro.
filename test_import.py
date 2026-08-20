import sys
import traceback

print("Testing direct import...")
try:
    from src.ingestion.trial_balance import TrialBalanceIngestor
    print("Success import:", TrialBalanceIngestor)
except Exception as e:
    print("Failed!")
    traceprint = traceback.format_exc()
    print(traceprint)
