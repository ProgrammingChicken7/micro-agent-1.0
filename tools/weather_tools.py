import requests
from datetime import datetime, timedelta
from config import get_tool_config

def call_weather_api(api_type, **params):
    """Unified interface function for calling WeatherAPI"""
    WEATHERAPI_KEY = get_tool_config('WEATHERAPI_KEY')
    BASE_URL = 'https://api.weatherapi.com/v1'
    endpoints = {
        'current': 'current.json', 'forecast': 'forecast.json', 'history': 'history.json',
        'marine': 'marine.json', 'future': 'future.json', 'astronomy': 'astronomy.json',
        'timezone': 'timezone.json', 'sports': 'sports.json', 'ip': 'ip.json',
        'search': 'search.json', 'alerts': 'alerts.json'
    }
    if api_type not in endpoints:
        raise ValueError(f"Invalid API type: {api_type}")
    
    location = params.get('location', 'auto:ip')
    request_params = {'key': WEATHERAPI_KEY, 'q': location, 'lang': params.get('lang', 'zh')}
    
    if api_type in ['forecast', 'marine']:
        request_params['days'] = min(params.get('days', 7), 14 if api_type == 'forecast' else 7)
    if api_type == 'forecast':
        request_params['alerts'] = params.get('alerts', 'yes')
        request_params['aqi'] = params.get('aqi', 'yes')
    if api_type == 'marine':
        request_params['tides'] = params.get('tides', 'yes')
    
    url = f"{BASE_URL}/{endpoints[api_type]}"
    try:
        response = requests.get(url, params=request_params, timeout=15)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"error": str(e)}
