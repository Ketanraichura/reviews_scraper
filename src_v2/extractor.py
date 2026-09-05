import re
import string

def normalize_whitespace(text):
    if not text:
        return None
    return re.sub(r'\s+', ' ', str(text)).strip()

def strip_noise(text):
    if not text:
        return None
    lines = text.split('\n')
    clean_lines = []
    
    stop_markers = [
        "are you human?",
        "This site is protected by reCAPTCHA",
        "Choose country",
        "Useful",
        "Share",
        "* * *",
        "![Image",
        "Trustpilot footer",
        "We collect device and interaction signals"
    ]
    
    for line in lines:
        stripped_line = line.strip()
        if not stripped_line:
            continue
            
        should_stop = False
        for marker in stop_markers:
            if marker.lower() in stripped_line.lower():
                should_stop = True
                break
        
        if should_stop:
            break
            
        clean_lines.append(stripped_line)
        
    return "\n".join(clean_lines).strip() if clean_lines else None

def extract_identity(content, expected_id, expected_text):
    if not content:
        return False, "NO_CONTENT"
        
    content_lower = content.lower()
    if expected_id and expected_id.lower() in content_lower:
        return True, "ID_MATCH"
        
    if expected_text:
        def norm_txt(t):
            t = str(t).lower().translate(str.maketrans('', '', string.punctuation))
            return re.sub(r'\s+', ' ', t).strip()
            
        norm_expected = norm_txt(expected_text)
        norm_content = norm_txt(content)
        
        if norm_expected and norm_expected in norm_content:
            return True, "EXACT_TEXT_MATCH"
            
        sentences = re.split(r'[.!?]+', str(expected_text))
        if sentences:
            longest_sentence = max(sentences, key=len).strip()
            norm_longest = norm_txt(longest_sentence)
            if len(norm_longest) > 20 and norm_longest in norm_content:
                return True, "STRONG_PARTIAL_TEXT_MATCH"
                
    return False, "INSUFFICIENT_EVIDENCE"

def extract_rating(content, title_meta=""):
    if title_meta:
        m = re.search(r'(\d)\s+stars?', title_meta, re.IGNORECASE)
        if m:
            return int(m.group(1))
            
    m = re.search(r'Rated\s+(\d)\s+out\s+of\s+5\s+stars', content, re.IGNORECASE)
    if m:
        return int(m.group(1))
        
    m = re.search(r'stars-(\d)\.svg', content, re.IGNORECASE)
    if m:
        return int(m.group(1))
        
    return None

def extract_review_date(content):
    lines = content.split('\n')
    date_pattern = re.compile(r"^(?:Updated\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4}$", re.IGNORECASE)
    
    found_rating = False
    for line in lines:
        line = line.strip()
        if "Rated " in line or "stars-" in line:
            found_rating = True
            continue
            
        if found_rating and date_pattern.match(line):
            return line
            
        if line.startswith("## ["):
            break
            
    return None

def extract_review_title_and_body(content):
    lines = content.split('\n')
    title = None
    body_lines = []
    in_body = False
    
    date_pattern = re.compile(r"^(?:Updated\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4}$", re.IGNORECASE)
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.startswith("## ["):
            m = re.search(r'## \[(.*?)\]\(', line)
            if m:
                title = m.group(1).strip()
            in_body = True
            continue
            
        if in_body:
            if line == "* * *" or "Useful" in line or "Share" in line:
                break
            if date_pattern.match(line) or "Date of experience:" in line:
                break
            if line.startswith("![Image") or line.startswith("[![Image"):
                continue
            body_lines.append(line)
            
    body = "\n".join(body_lines).strip() if body_lines else None
    
    raw_text = None
    if title and body:
        if title.lower() in body.lower():
            raw_text = body
        else:
            raw_text = f"{title}\n{body}"
    elif body:
        raw_text = body
    elif title:
        raw_text = title
        
    return raw_text

def extract_experience_date(content):
    m = re.search(r'Date of experience:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})', content, re.IGNORECASE)
    if m:
        return m.group(1).strip()
        
    lines = content.split('\n')
    in_body = False
    date_pattern = re.compile(r"^(?:Updated\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4}$", re.IGNORECASE)
    
    for line in lines:
        line = line.strip()
        if line.startswith("## ["):
            in_body = True
            continue
            
        if in_body:
            if date_pattern.match(line):
                return line
            if line == "* * *" or "Useful" in line or "Share" in line:
                break
                
    return None

def extract_company_reply_and_date(content):
    lines = content.split('\n')
    in_reply = False
    reply_lines = []
    reply_date = None
    
    date_pattern = re.compile(r"^(?:Updated\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4}$", re.IGNORECASE)
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if "Reply from" in line:
            in_reply = True
            continue
            
        if in_reply:
            if not reply_date and date_pattern.match(line):
                reply_date = line
                continue
                
            reply_lines.append(line)
            
    raw_reply = "\n".join(reply_lines).strip() if reply_lines else None
    clean_reply = strip_noise(raw_reply)
    
    return clean_reply, reply_date

def extract_all_fields(content, expected_id, expected_text, title_meta=""):
    is_match, reason = extract_identity(content, expected_id, expected_text)
    
    if not is_match:
        return {"Identity_Confirmed": False, "Reason": reason}
        
    reply_text, reply_date = extract_company_reply_and_date(content)
    
    return {
        "Identity_Confirmed": True,
        "Rating": extract_rating(content, title_meta),
        "Review_date": extract_review_date(content),
        "Raw_text": extract_review_title_and_body(content),
        "Order Date": extract_experience_date(content),
        "Reply Date": reply_date,
        "Support_reply": reply_text
    }
