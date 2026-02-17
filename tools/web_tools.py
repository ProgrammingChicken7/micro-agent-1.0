import requests
import re
import chardet
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
from rich.progress import Progress, SpinnerColumn, TextColumn
from config import get_tool_config

def raw_web_browser(url: str, show_details: bool = False):
    """Raw web browser: Get, parse and display detailed structure and content of a webpage"""
    def detect_encoding(response: requests.Response) -> str:
        content_type = response.headers.get('content-type', '')
        charset_match = re.search(r'charset=([^;]+)', content_type.lower())
        if charset_match:
            return charset_match.group(1).strip()
        html_snippet = response.content[:4096].decode('utf-8', errors='ignore')
        charset_patterns = [
            r'<meta[^>]*charset=["\s]*=["\s]*([^">\s]+)',
            r'<meta[^>]*content["\s]*=["\s]*[^"]*charset=([^">\s;]+)',
        ]
        for pattern in charset_patterns:
            match = re.search(pattern, html_snippet, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        detected = chardet.detect(response.content)
        if detected and detected['encoding'] and detected['confidence'] > 0.7:
            return detected['encoding']
        return 'utf-8'

    def get_html_content(target_url: str) -> str | None:
        if not urlparse(target_url).scheme:
            target_url = f'http://{target_url}'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
        }
        try:
            response = requests.get(target_url, headers=headers, timeout=15)
            response.raise_for_status()
            encoding = detect_encoding(response)
            try:
                return response.content.decode(encoding, errors='ignore')
            except:
                return response.content.decode('utf-8', errors='ignore')
        except:
            return None

    def clean_text(text: str | None) -> str:
        if not text: return ""
        return re.sub(r'\s+', ' ', text.strip())

    def extract_all_elements(html: str, base_url: str) -> dict:
        soup = BeautifulSoup(html, 'html.parser')
        for element in soup(["script", "style", "noscript", "link"]):
            element.decompose()
        elements = {}
        elements['title'] = clean_text(soup.title.string if soup.title else 'No title')
        meta_desc_tag = soup.find('meta', attrs={'name': 'description'})
        meta_keys_tag = soup.find('meta', attrs={'name': 'keywords'})
        elements['meta'] = {
            'description': meta_desc_tag['content'] if meta_desc_tag and meta_desc_tag.has_attr('content') else 'No description',
            'keywords': meta_keys_tag['content'] if meta_keys_tag and meta_keys_tag.has_attr('content') else 'No keywords'
        }
        main_content = soup.find('main') or soup.find('article') or soup.find('div', id=re.compile('content|main', re.I))
        body_text = main_content.get_text(separator='\n', strip=True) if main_content else soup.get_text(separator='\n', strip=True)
        elements['body'] = '\n'.join([line for line in body_text.split('\n') if line.strip()])
        elements['links'] = [{'url': urljoin(base_url, a['href']), 'text': clean_text(a.get_text()) or 'No text', 'title': clean_text(a.get('title', ''))} for a in soup.find_all('a', href=True)]
        elements['images'] = [{'src': urljoin(base_url, img['src']), 'alt': clean_text(img.get('alt', '')), 'title': clean_text(img.get('title', ''))} for img in soup.find_all('img', src=True)]
        return elements

    html = get_html_content(url)
    if not html: return {"error": f"Failed to retrieve content from {url}"}
    return extract_all_elements(html, url)

def perform_searchapi_search(query, engine="google", **kwargs):
    """
    Perform search using searchapi.io with support for all available engines.
    
    Supported engines:
    - Google family: google, google_images, google_videos, google_news, google_maps, 
                     google_shopping, google_flights, google_hotels, google_scholar, 
                     google_jobs, google_events, google_trends, google_finance, 
                     google_patents, google_lens, google_autocomplete
    - Other search engines: bing, bing_images, bing_videos, yahoo, baidu, yandex, 
                           duckduckgo, naver
    - E-commerce: amazon, shein, walmart, ebay
    - Social & Ads: youtube, meta_ad_library, linkedin_ad_library, tiktok_ads, 
                    facebook, instagram, tiktok_profile
    - Travel: airbnb, tripadvisor
    - Apps: google_play, apple_app_store
    
    Args:
        query: Search query string
        engine: Search engine to use (default: "google")
        **kwargs: Additional parameters specific to each engine (e.g., location, num, page)
    
    Returns:
        dict: JSON response from SearchAPI
    """
    SEARCHAPI_API_KEY = get_tool_config('SEARCHAPI_API_KEY')
    url = "https://www.searchapi.io/api/v1/search"
    
    # Build parameters
    params = {
        "engine": engine,
        "q": query,
        "api_key": SEARCHAPI_API_KEY
    }
    
    # Add optional parameters
    params.update(kwargs)
    
    try:
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"error": str(e)}

def search_google(query, **kwargs):
    """Search Google using SearchAPI"""
    return perform_searchapi_search(query, engine="google", **kwargs)

def search_google_images(query, **kwargs):
    """Search Google Images using SearchAPI"""
    return perform_searchapi_search(query, engine="google_images", **kwargs)

def search_google_videos(query, **kwargs):
    """Search Google Videos using SearchAPI"""
    return perform_searchapi_search(query, engine="google_videos", **kwargs)

def search_google_news(query, **kwargs):
    """Search Google News using SearchAPI"""
    return perform_searchapi_search(query, engine="google_news", **kwargs)

def search_google_maps(query, **kwargs):
    """Search Google Maps using SearchAPI"""
    return perform_searchapi_search(query, engine="google_maps", **kwargs)

def search_google_shopping(query, **kwargs):
    """Search Google Shopping using SearchAPI"""
    return perform_searchapi_search(query, engine="google_shopping", **kwargs)

def search_google_flights(query, **kwargs):
    """Search Google Flights using SearchAPI"""
    return perform_searchapi_search(query, engine="google_flights", **kwargs)

def search_google_hotels(query, **kwargs):
    """Search Google Hotels using SearchAPI"""
    return perform_searchapi_search(query, engine="google_hotels", **kwargs)

def search_google_scholar(query, **kwargs):
    """Search Google Scholar using SearchAPI"""
    return perform_searchapi_search(query, engine="google_scholar", **kwargs)

def search_google_jobs(query, **kwargs):
    """Search Google Jobs using SearchAPI"""
    return perform_searchapi_search(query, engine="google_jobs", **kwargs)

def search_google_events(query, **kwargs):
    """Search Google Events using SearchAPI"""
    return perform_searchapi_search(query, engine="google_events", **kwargs)

def search_google_trends(query, **kwargs):
    """Search Google Trends using SearchAPI"""
    return perform_searchapi_search(query, engine="google_trends", **kwargs)

def search_google_finance(query, **kwargs):
    """Search Google Finance using SearchAPI"""
    return perform_searchapi_search(query, engine="google_finance", **kwargs)

def search_google_patents(query, **kwargs):
    """Search Google Patents using SearchAPI"""
    return perform_searchapi_search(query, engine="google_patents", **kwargs)

def search_google_lens(query, **kwargs):
    """Search Google Lens using SearchAPI"""
    return perform_searchapi_search(query, engine="google_lens", **kwargs)

def search_google_autocomplete(query, **kwargs):
    """Get Google Autocomplete suggestions using SearchAPI"""
    return perform_searchapi_search(query, engine="google_autocomplete", **kwargs)

def search_bing(query, **kwargs):
    """Search Bing using SearchAPI"""
    return perform_searchapi_search(query, engine="bing", **kwargs)

def search_bing_images(query, **kwargs):
    """Search Bing Images using SearchAPI"""
    return perform_searchapi_search(query, engine="bing_images", **kwargs)

def search_bing_videos(query, **kwargs):
    """Search Bing Videos using SearchAPI"""
    return perform_searchapi_search(query, engine="bing_videos", **kwargs)

def search_yahoo(query, **kwargs):
    """Search Yahoo using SearchAPI"""
    return perform_searchapi_search(query, engine="yahoo", **kwargs)

def search_baidu(query, **kwargs):
    """Search Baidu using SearchAPI"""
    return perform_searchapi_search(query, engine="baidu", **kwargs)

def search_yandex(query, **kwargs):
    """Search Yandex using SearchAPI"""
    return perform_searchapi_search(query, engine="yandex", **kwargs)

def search_duckduckgo(query, **kwargs):
    """Search DuckDuckGo using SearchAPI"""
    return perform_searchapi_search(query, engine="duckduckgo", **kwargs)

def search_naver(query, **kwargs):
    """Search Naver using SearchAPI"""
    return perform_searchapi_search(query, engine="naver", **kwargs)

def search_amazon(query, **kwargs):
    """Search Amazon products using SearchAPI"""
    return perform_searchapi_search(query, engine="amazon", **kwargs)

def search_shein(query, **kwargs):
    """Search Shein products using SearchAPI"""
    return perform_searchapi_search(query, engine="shein", **kwargs)

def search_walmart(query, **kwargs):
    """Search Walmart products using SearchAPI"""
    return perform_searchapi_search(query, engine="walmart", **kwargs)

def search_ebay(query, **kwargs):
    """Search eBay products using SearchAPI"""
    return perform_searchapi_search(query, engine="ebay", **kwargs)

def search_youtube(query, **kwargs):
    """Search YouTube videos using SearchAPI"""
    return perform_searchapi_search(query, engine="youtube", **kwargs)

def search_google_play(query, **kwargs):
    """Search Google Play Store using SearchAPI"""
    return perform_searchapi_search(query, engine="google_play", **kwargs)

def search_apple_app_store(query, **kwargs):
    """Search Apple App Store using SearchAPI"""
    return perform_searchapi_search(query, engine="apple_app_store", **kwargs)

def search_airbnb(query, **kwargs):
    """Search Airbnb listings using SearchAPI"""
    return perform_searchapi_search(query, engine="airbnb", **kwargs)

def search_tripadvisor(query, **kwargs):
    """Search TripAdvisor using SearchAPI"""
    return perform_searchapi_search(query, engine="tripadvisor", **kwargs)

def search_meta_ad_library(query, **kwargs):
    """Search Meta (Facebook) Ad Library using SearchAPI"""
    return perform_searchapi_search(query, engine="meta_ad_library", **kwargs)

def search_linkedin_ad_library(query, **kwargs):
    """Search LinkedIn Ad Library using SearchAPI"""
    return perform_searchapi_search(query, engine="linkedin_ad_library", **kwargs)

def search_tiktok_ads(query, **kwargs):
    """Search TikTok Ads Library using SearchAPI"""
    return perform_searchapi_search(query, engine="tiktok_ads", **kwargs)

def search_tiktok_profile(query, **kwargs):
    """Get TikTok profile information using SearchAPI"""
    return perform_searchapi_search(query, engine="tiktok_profile", **kwargs)

def search_facebook_page(query, **kwargs):
    """Get Facebook business page information using SearchAPI"""
    return perform_searchapi_search(query, engine="facebook", **kwargs)

def search_instagram_profile(query, **kwargs):
    """Get Instagram profile information using SearchAPI"""
    return perform_searchapi_search(query, engine="instagram", **kwargs)

# Legacy Google and Bing search functions removed (API keys invalid)
# Use SearchAPI (perform_searchapi_search) or Scira AI (search_scira) instead

def search_scira(query, agent="search", timeout=30, **kwargs):
    """
    Perform search using Scira AI with support for 4 specialized agents.
    
    Supported agents:
    - search: General web search with AI summary and sources
    - people: Find and enrich people profiles (LinkedIn, GitHub, X, etc.)
    - xsearch: Real-time X (Twitter) search, optionally target specific user
    - reddit: Search Reddit discussions with AI summary
    
    Args:
        query: Search query string
        agent: Agent type (default: "search")
        timeout: Request timeout in seconds (default: 30)
        **kwargs: Additional parameters (e.g., username for xsearch, messages for search)
    
    Returns:
        dict: JSON response from Scira API with text, sources, and results
    """
    SCIRA_API_KEY = get_tool_config('SCIRA_API_KEY')
    
    # Agent endpoint mapping
    agent_endpoints = {
        "search": "/api/search",
        "people": "/api/people",
        "xsearch": "/api/xsearch",
        "reddit": "/api/reddit"
    }
    
    if agent not in agent_endpoints:
        return {"error": f"Invalid agent type: {agent}. Must be one of: {', '.join(agent_endpoints.keys())}"}
    
    url = f"https://api.scira.ai{agent_endpoints[agent]}"
    headers = {
        "x-api-key": SCIRA_API_KEY,
        "Content-Type": "application/json"
    }
    
    # Build request data based on agent type
    if agent == "search":
        # Search agent uses messages format
        messages = kwargs.get('messages', [{"role": "user", "content": query}])
        data = {"messages": messages}
    elif agent == "people":
        # People agent uses query format
        data = {"query": query}
    elif agent == "xsearch":
        # X agent supports optional username
        data = {"query": query}
        if 'username' in kwargs:
            data['username'] = kwargs['username']
    elif agent == "reddit":
        # Reddit agent uses query format
        data = {"query": query}
    
    try:
        response = requests.post(url, headers=headers, json=data, timeout=timeout)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as e:
        status_code = e.response.status_code
        error_msg = e.response.json().get('error', str(e)) if e.response.text else str(e)
        return {"error": f"HTTP {status_code}: {error_msg}"}
    except Exception as e:
        return {"error": str(e)}

def search_scira_web(query, messages=None, timeout=30):
    """Search the web using Scira AI Search Agent"""
    kwargs = {}
    if messages:
        kwargs['messages'] = messages
    return search_scira(query, agent="search", timeout=timeout, **kwargs)

def search_scira_people(query, timeout=30):
    """Find and enrich people profiles using Scira AI People Agent"""
    return search_scira(query, agent="people", timeout=timeout)

def search_scira_x(query, username=None, timeout=30):
    """Search X (Twitter) in real-time using Scira AI X Agent"""
    kwargs = {}
    if username:
        kwargs['username'] = username
    return search_scira(query, agent="xsearch", timeout=timeout, **kwargs)

def search_scira_reddit(query, timeout=30):
    """Search Reddit discussions using Scira AI Reddit Agent"""
    return search_scira(query, agent="reddit", timeout=timeout)

def get_ip_geolocation(ip=""):
    """Get IP geolocation information"""
    IPGEOLOCATION_API_KEY = get_tool_config('IPGEOLOCATION_API_KEY')
    url = f"https://api.ipgeolocation.io/ipgeo?apiKey={IPGEOLOCATION_API_KEY}"
    if ip: url += f"&ip={ip}"
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"error": str(e)}
