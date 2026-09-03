# Trustpilot Review Verifier & Automated Audit Pipeline

An enterprise-grade, defensive audit and verification engine designed to validate, correct, and cross-reference Trustpilot review datasets against live web evidence. Built with atomic checkpointing, anti-hallucination safeguards, zero-data-loss guarantees, and comprehensive audit trail logging.

---

## 📑 Table of Contents

1. [System Architecture](#system-architecture)
2. [Core Verification & Safeguard Engine](#core-verification--safeguard-engine)
3. [The 8-Column Audit Trail](#the-8-column-audit-trail)
4. [Confidence Scoring & Governance](#confidence-scoring--governance)
5. [Adaptive Rate Limiting & Resiliency](#adaptive-rate-limiting--resiliency)
6. [Benchmark Analysis: Local vs. Cloud (Codespaces)](#benchmark-analysis-local-vs-cloud-codespaces)
7. [Repository & File Structure](#repository--file-structure)
8. [Installation & Setup](#installation--setup)
9. [Operational Runbook (How to Run)](#operational-runbook-how-to-run)
10. [Troubleshooting & FAQs](#troubleshooting--faqs)

---

## 🏛 System Architecture

The pipeline processes Excel workbooks (`.xlsm` / `.xlsx`) containing historical Trustpilot review records, queries live web representations of each review, verifies review identity, compares target fields, applies strict corrections where discrepancies exist, and writes out a fully audited workbook while preserving all existing VBA macros.

```
┌────────────────────────┐
│  Input: LG_corrected   │
│   (VBA .xlsm file)     │
└───────────┬────────────┘
            │
            ▼
┌────────────────────────────────────────────────────────┐
│             TrustpilotVerifier Engine                  │
│                                                        │
│  1. Jina Reader Acquisition Layer (https://r.jina.ai)  │
│  2. Identity Matching (Review ID & Normalized Text)    │
│  3. Defensive Field Extraction & Safeguards            │
│  4. Discrepancy Detection & Field Normalization        │
│  5. Adaptive Concurrency & Exponential Backoff        │
└───────────┬────────────────────────────────────────────┘
            │
      ┌─────┴─────────────────────────┐
      ▼                               ▼
┌───────────────────────────┐   ┌───────────────────────────┐
│ checkpoint_prod.json      │   │ LG_audited_final.xlsm     │
│ (Atomic State Persistence)│   │ (Preserved VBA + Audit)   │
└───────────────────────────┘   └───────────────────────────┘
```

### Key Architectural Pillars
- **Headless-Free Retrieval:** Queries `https://r.jina.ai/{trustpilot_url}` to convert raw DOM into structured Markdown, avoiding the operational overhead, memory leakage, and browser fingerprint traps of heavy headless browser automation.
- **Defensive Data Governance ("Do No Harm"):** The pipeline will **never** overwrite an existing field with a truncated string, a partial extraction, or an unverified match.
- **VBA Macro Preservation:** Output workbooks are written using `openpyxl` with `keep_vba=True` to guarantee that all embedded macros, buttons, and VBA modules remain functional.
- **Zero-Loss Atomic Checkpointing:** Every batch of processed rows flushes to a persistent JSON state file. If a network disruption or SIGINT occurs, the engine resumes from the exact row where it stopped without repeating completed rows.

---

## 🛡 Core Verification & Safeguard Engine

### 1. Two-Tier Identity Verification
Before any comparison or correction can occur, the pipeline must mathematically prove that the downloaded page corresponds to the exact review in that row:
1. **Tier 1 (Hexadecimal Review ID):** Extracts the 24-character hexadecimal ID from the review URL (e.g., `/reviews/64b58e72f10b7418...`) and verifies its presence in the fetched Markdown.
2. **Tier 2 (Normalized Content Equality):** If the review ID is absent (such as on company overview pages), the engine normalizes the expected review text (lowercase, stripped punctuation, normalized whitespace) and checks for strict substring containment within the normalized document text.

If both checks fail, the status is flagged as `MULTIPLE_POSSIBLE_MATCHES` and **no changes are made**.

### 2. Field Extraction & QA Safeguards
The engine monitors and compares 5 core fields:

| Field | Source Column | Safeguards & Parsing Logic |
| :--- | :---: | :--- |
| **`Raw_text`** | Column 3 (`C`) | **Anti-Truncation Guard:** Rejects Trustpilot values ending in `...` or `…`.<br>**Length Ratio Guard:** If extracted length is `< 50%` of original length, the extraction is rejected to prevent accidental truncation. |
| **`Rating`** | Column 4 (`D`) | Extracts 1–5 integer rating from Trustpilot star indicators (`Rated X out of 5 stars`). |
| **`Review_date`** | Column 5 (`E`) | ISO 8601 UTC timestamp parsed and converted to normalized calendar date format (`YYYY-MM-DD`). |
| **`Reply_date`** | Column 6 (`F`) | Company response date parsed and standardized. |
| **`Support_reply`** | Column 7 (`G`) | Full company reply text extracted and stripped of whitespace noise. |

---

## 📊 The 8-Column Audit Trail

When verification completes, 8 standardized audit columns are appended directly to the right of the existing table headers in the output workbook:

```
[ Col 1..N: Original Data ] | [ Col N+1..N+8: Audit Trail ]
```

1. **`Verification_Status`**: High-level classification code:
   - `VERIFIED_MATCH`: Identity confirmed; all fields match Trustpilot perfectly.
   - `VERIFIED_AND_CORRECTED`: Identity confirmed; one or more fields differed and were safely updated.
   - `MULTIPLE_POSSIBLE_MATCHES`: Page retrieved, but identity could not be verified with 100% certainty.
   - `REVIEW_NOT_FOUND`: URL returned a confirmed HTTP 404 (review deleted or URL broken).
   - `ACCESS_LIMITED`: Upstream challenge screen (Cloudflare/Trustpilot verification).
   - `RATE_LIMITED`: Upstream HTTP 429 received after backoff retries.
   - `SOURCE_DATA_INSUFFICIENT`: Missing target URL or invalid source row.
2. **`Field_Affected`**: Names of modified fields (`Raw_text`, `Rating`, `Review_date`, `Reply_date`, `Support_reply`, or `None`). Multiple fields are pipe-delimited (`Rating | Review_date`).
3. **`Original_Value`**: The snapshot of the value before correction.
4. **`Trustpilot_Value`**: The live value extracted from Trustpilot.
5. **`Correction_Made`**: Explicit flag: `Yes` or `No`.
6. **`Evidence_Source`**: Origin of proof (`Jina Markdown Extract`, `Jina Markdown`, `HTTP Response`).
7. **`Confidence`**: The certainty grade (`High` or `Low`).
8. **`Reason_Action_Taken`**: Human-readable explanation of why the action was taken.

---

## 🎯 Confidence Scoring & Governance

The audit trail includes an explicit **`Confidence`** classification (`High` vs. `Low`).

### Why Log "Low" Confidence When No Changes Are Made?
If an automated system leaves a row untouched, there are two fundamentally different reasons:
1. **High Confidence (`VERIFIED_MATCH`):** The review was verified and the data was already 100% correct. **No action needed.**
2. **Low Confidence (`MULTIPLE_POSSIBLE_MATCHES`):** Content was fetched, but the script could not prove which review belonged to this row. The row was left untouched **solely because the script refused to guess**. **Action may be needed.**

Logging `Confidence: Low` allows an analyst to filter by `Confidence = Low` in Excel and inspect edge cases within seconds.

| Confidence | Typical Status | Operational Meaning | Action Taken |
| :---: | :--- | :--- | :---: |
| **`High`** | `VERIFIED_MATCH`<br>`VERIFIED_AND_CORRECTED` | Identity and field extracts are verified with 100% certainty. | Safe updates applied. |
| **`High`** | `REVIEW_NOT_FOUND`<br>`RATE_LIMITED` | HTTP status code is definitive (e.g., confirmed 404 or 429). | Flagged for network retry. |
| **`Low`** | `MULTIPLE_POSSIBLE_MATCHES` | Page retrieved, but identity ambiguous. | **Zero changes made.** Flagged for manual review. |

---

## ⚡ Adaptive Rate Limiting & Resiliency

To balance execution speed with upstream rate-limit thresholds, `TrustpilotVerifier` implements a dynamic concurrency controller:

- **Thread Pool Management:** Bounded concurrency using Python's `concurrent.futures.ThreadPoolExecutor`.
- **Scaling Up on Success:** After **20 consecutive successful requests**, active concurrency increments by 1 (up to `max_concurrency`).
- **Immediate Circuit Breaker on Throttling:** Upon receiving **2 consecutive errors** (`RATE_LIMITED`, `ACCESS_LIMITED`, or network failures):
  1. Concurrency decreases by 2 (down to a minimum of 1).
  2. The thread pool immediately pauses with a **5-second cooldown** (`time.sleep(5)`).
  3. Individual workers apply exponential backoff with jitter (`2^attempt + random(0, 1)`).

---

## 🔬 Benchmark Analysis: Local vs. Cloud (Codespaces)

An isolated benchmark was executed across 100-row sample batches to determine whether migrating execution to a GitHub Codespaces cloud instance (4-core, 16GB RAM, Azure Datacenter Egress `4.240.39.203`) yields higher sustainable throughput than a local development machine.

### Benchmark Results Table

| Environment | Initial Concurrency | Settled Concurrency | Sustainable Req/s | 429 Rate | Access-Limited Rate | Avg Latency | p95 Latency | Wall Clock (100 rows) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Local Environment** | ~1–3 | 1–2 | **~0.30** | Known High | Known | ~2.5–3.5s | ~7.5s | ~330s |
| **Codespaces (Tier 4)** | 10 | 2 | **0.39** | **14.0%** (14/100) | **2.0%** (2/100) | 0.88s | 4.48s | 215.07s |
| **Codespaces (Tier 3)** | 5 | 1–2 | **0.33** | **11.0%** (11/100) | **1.0%** (1/100) | 0.60s | 1.52s | 270.23s |

### Core Findings & Conclusions
1. **The Bottleneck is Upstream:** While Codespaces provides significantly faster round-trip network latency (0.60s vs. 3.0s+), Jina Reader and Trustpilot enforce a strict rate limit per egress IP.
2. **Instant Burst vs. Sustainable Ceiling:** At 10 workers, Codespaces achieves an initial burst of **12.4 req/s**. However, after ~25 requests, Jina Reader's 429 throttle triggers, causing the verifier's backoff to drop active concurrency to 1–2 workers.
3. **Migration Recommendation:** **Do not migrate production to Codespaces.** The sustainable throughput ceiling (~0.33–0.39 req/s) is bounded by Jina's upstream API limits, making migration unnecessary.

---

## 📁 Repository & File Structure

```
reviews_scraper/
├── LG_corrected.xlsm               # Master input workbook (2,198 rows, macros preserved)
├── LG_audited_final.xlsm           # Fully audited production output with 8 audit columns
├── checkpoint_prod.json            # Persistent atomic state for production run
├── checkpoint_qa.json              # Persistent state for 100-row QA stress tests
├── run_production.py               # Production execution entry point (full 2,198 rows)
├── run_stress_test.py              # Rapid 100-row concurrent verification test
├── run_benchmark.py                # Multi-tier concurrency benchmark suite
├── benchmark_output/               # Isolated outputs and checkpoints from benchmark runs
│   ├── LG_benchmark_c10.xlsm
│   ├── checkpoint_c10.json
│   ├── LG_benchmark_c5.xlsm
│   └── checkpoint_c5.json
└── src/
    ├── __init__.py
    ├── reader.py                   # Acquisition layer client for Jina Reader API
    └── verifier.py                 # Core verification, extraction, and audit engine
```

---

## ⚙️ Installation & Setup

### Prerequisites
- Python 3.10+ (tested on Python 3.12 and 3.13)
- `pip` package manager

### 1. Clone & Navigate
```bash
git clone https://github.com/Ketanraichura/reviews_scraper-copy.git
cd reviews_scraper-copy
```

### 2. Install Dependencies
```bash
pip install openpyxl requests
```

---

## 🚀 Operational Runbook (How to Run)

### 1. Run a Fast QA Stress Test (100 Rows)
Validates verification integrity and rate limiting on a 100-row subset:
```bash
python3 run_stress_test.py
```
- Output saved to: `LG_qa_test.xlsm`
- State saved to: `checkpoint_qa.json`

### 2. Run the Full Production Run (2,198 Rows)
Executes verification across the entire dataset with automatic state resumption:
```bash
python3 run_production.py
```
- Output saved to: `LG_audited_final.xlsm`
- State saved to: `checkpoint_prod.json`

### 3. Run Concurrency Benchmarks
To test throughput across different worker bounds (e.g. 5 vs. 10 workers):
```bash
python3 run_benchmark.py
```
- All artifacts are written to `benchmark_output/` without touching production files.

### 4. Resuming an Interrupted Run
The script automatically detects existing checkpoint files. If a run is stopped via `Ctrl+C` or a network disconnection:
1. Re-run `python3 run_production.py`.
2. The engine logs:
   ```
   Total target rows: 2198
   Already completed: 760
   Pending processing: 1438
   ```
3. Execution resumes seamlessly from row 761.

---

## ❓ Troubleshooting & FAQs

#### Q: How do I know if an extraction was truncated?
**A:** The engine inspects every extracted `Raw_text` value. If it terminates with ellipsis (`...` or `…`), or if its length is less than 50% of the original string, the update is rejected and the original text is preserved.

#### Q: What does HTTP 429 mean in the audit trail?
**A:** HTTP 429 indicates that Jina Reader's API or Trustpilot rate-limited the connection. The engine automatically retries with exponential backoff before logging `RATE_LIMITED`.

#### Q: Are Excel VBA macros safe?
**A:** Yes. The verifier opens and saves the workbook using `openpyxl.load_workbook(..., keep_vba=True)`. All macros, button assignments, and underlying VBA project streams remain intact.
