import os
import json
import time
import random
import requests

CACHE_DIR = "cache/raw_reviews"

def ensure_cache_dir():
    os.makedirs(CACHE_DIR, exist_ok=True)

def get_cache_path(review_id):
    if not review_id: return None
    return os.path.join(CACHE_DIR, f"{review_id}.json")

def load_cached(review_id):
    path = get_cache_path(review_id)
    if path and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def save_cache(review_id, data):
    ensure_cache_dir()
    path = get_cache_path(review_id)
    if path:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

def fetch_jina(url, review_id, max_retries=4):
    cached = load_cached(review_id)
    if cached:
        return cached, True # True means loaded from cache
        
    jina_url = f"https://r.jina.ai/{url}"
    headers = {"Accept": "application/json"}
    
    for attempt in range(max_retries):
        try:
            resp = requests.get(jina_url, headers=headers, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                title = data.get("data", {}).get("title", "")
                content = data.get("data", {}).get("content", "")
                
                if "Verifying your connection" in title or "challenge.js" in content or "Verifying Connection" in title:
                    # Cloudflare challenge
                    time.sleep((2 ** attempt) + random.uniform(0.5, 2.0))
                    continue
                    
                if "404" in title or "Page not found" in title:
                    return {"status": "REVIEW_NOT_FOUND", "data": None}, False
                    
                res = {"status": "SUCCESS", "data": data.get("data")}
                save_cache(review_id, res)
                return res, False
                
            elif resp.status_code == 429:
                time.sleep((2 ** attempt) * 2 + random.uniform(1.0, 3.0))
                continue
            else:
                time.sleep(2)
                continue
                
        except Exception as e:
            time.sleep(2)
            continue
            
    return {"status": "RATE_LIMITED_OR_BLOCKED", "data": None}, False
