from dateutil import parser as date_parser
import string
import re

def is_empty(val):
    if val is None: return True
    s = str(val).strip()
    return s in ["", "-", "None", "N/A", "nan"]

def are_dates_equal(val1, val2):
    try:
        s1 = str(val1).replace("Updated ", "").strip()
        if hasattr(val1, "date"): d1 = val1.date()
        else: d1 = date_parser.parse(s1).date()
        
        s2 = str(val2).replace("Updated ", "").strip()
        if hasattr(val2, "date"): d2 = val2.date()
        else: d2 = date_parser.parse(s2).date()
        
        return d1 == d2
    except Exception:
        return False

def compare_date(orig_val, tp_val):
    orig_empty = is_empty(orig_val)
    tp_empty = is_empty(tp_val)
    
    if orig_empty and tp_empty:
        return "MATCH", orig_val, False
        
    if orig_empty and not tp_empty:
        return "RECOVERED", tp_val, False
        
    if not orig_empty and tp_empty:
        # Trustpilot didn't have it, keep original, no discrepancy unless we want to flag deletion
        return "MATCH", orig_val, False
        
    if are_dates_equal(orig_val, tp_val):
        # Normalization occurred if strings differ but calendar day is same
        norm_applied = str(orig_val).strip() != str(tp_val).strip()
        return "MATCH", orig_val, norm_applied
        
    return "DISCREPANCY", tp_val, False

def compare_rating(orig_val, tp_val):
    orig_empty = is_empty(orig_val)
    tp_empty = is_empty(tp_val)
    
    if orig_empty and tp_empty:
        return "MATCH", orig_val, False
        
    if orig_empty and not tp_empty:
        return "RECOVERED", tp_val, False
        
    if not orig_empty and tp_empty:
        return "MATCH", orig_val, False
        
    try:
        o_float = float(orig_val)
        t_float = float(tp_val)
        if o_float == t_float:
            return "MATCH", orig_val, False
        else:
            return "DISCREPANCY", tp_val, False
    except ValueError:
        return "DISCREPANCY", tp_val, False

def normalize_text_for_compare(t):
    if not t: return ""
    t = str(t).replace('\r\n', '\n').strip()
    # Remove smart quotes, em-dashes for robust comparison
    t = t.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
    t = t.replace('—', '-').replace('–', '-')
    # Condense whitespace
    t = re.sub(r'\s+', ' ', t)
    # Remove trailing punctuation safely
    t = t.rstrip(string.punctuation)
    return t.lower()

def compare_text(orig_val, tp_val):
    orig_empty = is_empty(orig_val)
    tp_empty = is_empty(tp_val)
    
    if orig_empty and tp_empty:
        return "MATCH", orig_val, False
        
    if orig_empty and not tp_empty:
        return "RECOVERED", tp_val, False
        
    if not orig_empty and tp_empty:
        return "MATCH", orig_val, False
        
    norm_orig = normalize_text_for_compare(orig_val)
    norm_tp = normalize_text_for_compare(tp_val)
    
    if norm_orig == norm_tp:
        norm_applied = str(orig_val).strip() != str(tp_val).strip()
        return "MATCH", orig_val, norm_applied
        
    # Safe fallback if extracted text is suspiciously truncated
    if len(str(tp_val)) < (len(str(orig_val)) * 0.5):
        # We might have missed part of the text, don't overwrite blindly
        return "MATCH", orig_val, False
        
    return "DISCREPANCY", tp_val, False

def compare_all(row_data, tp_data):
    results = {}
    
    # Text fields
    for field in ["Raw_text", "Support_reply"]:
        verdict, final_val, norm = compare_text(row_data.get(field), tp_data.get(field))
        results[field] = {"verdict": verdict, "final_val": final_val, "norm": norm, "orig_val": row_data.get(field), "tp_val": tp_data.get(field)}
        
    # Date fields
    for field in ["Review_date", "Order Date", "Reply Date"]:
        verdict, final_val, norm = compare_date(row_data.get(field), tp_data.get(field))
        results[field] = {"verdict": verdict, "final_val": final_val, "norm": norm, "orig_val": row_data.get(field), "tp_val": tp_data.get(field)}
        
    # Rating
    verdict, final_val, norm = compare_rating(row_data.get("Rating"), tp_data.get("Rating"))
    results["Rating"] = {"verdict": verdict, "final_val": final_val, "norm": norm, "orig_val": row_data.get("Rating"), "tp_val": tp_data.get("Rating")}
    
    return results
