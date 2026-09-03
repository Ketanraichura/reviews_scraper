import json
import time
import os
from urllib.parse import urlparse, urljoin
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

BUSINESS_PROFILES = {
    "low_volume": "anthropic.com",      # < 200 reviews
    "medium_volume": "replit.com",      # ~500-2000 reviews
    "high_volume": "shopify.com"        # 10,000+ reviews
}

TARGET_PAGES = [9, 10, 11, 12]
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "output", "experiment_results.json")

def analyze_page(context, url, domain, page_num):
    print(f"\n[{domain}] Testing Page {page_num} -> {url}")
    
    page = context.new_page()
    
    # Store relevant API requests
    api_requests = []
    
    def handle_response(response):
        # Record JSON data routes or graphql requests to see if Trustpilot loads reviews via API
        req_url = response.url
        if "trustpilot.com" in req_url and ("/_next/data/" in req_url or "/api/" in req_url or "graphql" in req_url.lower()):
            api_requests.append({
                "url": req_url,
                "status": response.status,
                "content_type": response.headers.get("content-type", "")
            })

    page.on("response", handle_response)
    
    result = {
        "requested_url": url,
        "final_url": None,
        "http_status": None,
        "page_title": None,
        "waf_challenge_detected": False,
        "number_of_reviews_extracted": 0,
        "unique_review_ids": [],
        "duplicate_review_ids_detected": False,
        "pagination_metadata": {},
        "api_requests": [],
        "content_type": None,
        "error": None
    }
    
    try:
        response = page.goto(url, wait_until="domcontentloaded", timeout=30000)
        
        result["final_url"] = page.url
        if response:
            result["http_status"] = response.status
            result["content_type"] = response.headers.get("content-type")
        
        # Give it a second to see if there's a quick JS challenge or rendering
        page.wait_for_timeout(2000)
        
        result["page_title"] = page.title()
        html_content = page.content()
        
        # Check for AWS WAF or Cloudflare Interstitial
        if "Verifying your connection" in html_content or "challenge.js" in html_content or "awswaf.com" in html_content:
            result["waf_challenge_detected"] = True
            
        # Extract reviews using BeautifulSoup
        soup = BeautifulSoup(html_content, "html.parser")
        
        # Trustpilot usually stores reviews in `<article>` or elements with specific classes/data-attributes
        review_elements = soup.find_all("article")
        
        review_ids = []
        for el in review_elements:
            # Often the ID is on the article or an inner element
            r_id = el.get("id") or el.get("data-review-id")
            # Or look for an inner child with ID if not on article
            if not r_id:
                # Common fallback
                child_with_id = el.find(id=True)
                if child_with_id:
                    r_id = child_with_id.get("id")
            if r_id:
                review_ids.append(r_id)
                
        # Some Trustpilot layouts use '__NEXT_DATA__' for the page state
        next_data_script = soup.find("script", id="__NEXT_DATA__")
        if next_data_script:
            result["pagination_metadata"]["has_next_data_json"] = True
            try:
                data = json.loads(next_data_script.string)
                # Try to find reviews in the json if they aren't in the DOM
                if not review_ids:
                    props = data.get("props", {}).get("pageProps", {})
                    reviews = props.get("reviews", [])
                    if reviews:
                        review_ids = [r.get("id") for r in reviews if r.get("id")]
            except Exception:
                pass
                
        # Look for next page links
        next_link = soup.find("a", attrs={"data-pagination-button-next-link": "true"}) or soup.find("link", rel="next")
        if next_link:
            result["pagination_metadata"]["next_url"] = next_link.get("href")
            
        # If still no review IDs, try basic text search for common patterns to estimate counts
        if not review_ids:
            # Generic fallback if specific selectors fail but content is there
            result["number_of_reviews_extracted"] = len(soup.find_all(attrs={"data-review-content": True}))
        
        result["unique_review_ids"] = list(set(review_ids))
        result["number_of_reviews_extracted"] = len(review_ids)
        
        if len(review_ids) > len(result["unique_review_ids"]):
            result["duplicate_review_ids_detected"] = True
            
    except Exception as e:
        result["error"] = str(e)
        
    result["api_requests"] = api_requests
    page.close()
    
    return result

def main():
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    
    results = {}
    
    with sync_playwright() as p:
        # Launching a normal Chromium browser
        browser = p.chromium.launch(headless=True)
        # Using a typical user agent
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        )
        
        for volume_type, domain in BUSINESS_PROFILES.items():
            results[domain] = {
                "volume_type": volume_type,
                "pages": {}
            }
            
            for page_num in TARGET_PAGES:
                url = f"https://www.trustpilot.com/review/{domain}?page={page_num}"
                page_data = analyze_page(context, url, domain, page_num)
                results[domain]["pages"][page_num] = page_data
                
                # Small pause to be polite and avoid basic rate limiting (though WAF might trigger regardless)
                time.sleep(2)
                
        browser.close()
        
    with open(OUTPUT_FILE, "w") as f:
        json.dump(results, f, indent=2)
        
    print(f"\nExperiment complete. Results saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
