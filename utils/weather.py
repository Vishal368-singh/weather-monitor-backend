import os
from dotenv import load_dotenv
import requests
from psycopg2.extras import RealDictCursor
import base64
from utils.db import get_db_conn, release_db_conn
load_dotenv()

weatherAPIKey = os.environ.get("weatherAPIKey")
visualCrossAPIKey = os.environ.get("visualCrossAPIKey")
weather_icon_base_url = os.environ.get("WEATHER_ICON_BASEURL")

# CURRENT_MAP_WAPI_VCAPI = {
#     "datetimeEpoch": "last_updated_epoch",
#     "datetime": "last_updated",
#     "temp": "temp_c",
#     "conditions": "condition",
#     "windspeed": "wind_kph",
#     "winddir": "wind_dir",
#     "pressure": "pressure_mb",
#     "precip": "precip_mm",
#     "humidity": "humidity",
#     "cloudcover": "cloud",
#     "feelslike": "feelslike_c",
#     "visibility": "vis_km",
#     "uvindex": "uv"
# }

# DAYFORECAST_MAP_WAPI_VCAPI = {
#     "datetime": "date",
#     "datetimeEpoch": "date_epoch",
#     "tempmax": "maxtemp_c",
#     "tempmin": "mintemp_c",
#     "windspeed": "maxwind_kph",
#     "precip": "totalprecip_mm",
#     "snow": "totalsnow_cm",
#     "visibility": "avgvis_km",
#     "humidity": "avghumidity",
#     "precipprob": "daily_chance_of_rain",
#     "conditions": "condition",
#     "uvindex": "uv",
#     "sunrise": "sunrise",
#     "sunset": "sunset",
#     "moonphase": "moon_phase"
# }

# DAYFORECAST_MAP_WAPI_VCAPI = {
#     "datetimeEpoch": "time_epoch",
#     "datetime": "time",
#     "temp": "temp_c",
#     "conditions": "condition",
#     "windspeed": "wind_kph",
#     "winddir": "wind_dir",
#     "windgust": "gust_kph",
#     "pressure": "pressure_mb",
#     "precip": "precip_mm",
#     "snow": "snow_cm",
#     "humidity": "humidity",
#     "cloudcover": "cloud",
#     "feelslike": "feelslike_c",
#     "dew": "dewpoint_c",
#     "precipprob": "chance_of_rain",
#     "visibility": "vis_km",
#     "uvindex": "uv"
# }

VC_TO_WEATHERAPI_MAP = {
    # Date level
    "datetime": "date",
    "datetimeEpoch": "date_epoch",

    # Daily
    "tempmax": "maxtemp_c",
    "tempmin": "mintemp_c",
    "windspeed": "maxwind_kph",
    "precip": "totalprecip_mm",
    "snow": "totalsnow_cm",
    "visibility": "avgvis_km",
    "humidity": "avghumidity",
    "precipprob": "daily_chance_of_rain",
    "uvindex": "uv",
    "moonphase": "moon_phase",

    # Hourly
    "temp": "temp_c",
    "winddir": "wind_dir",
    "pressure": "pressure_mb",
    "cloudcover": "cloud",
    "feelslike": "feelslike_c",
    "dew": "dewpoint_c",
    "windgust": "gust_kph",
    "precip": "precip_mm",
    "precipprob": "chance_of_rain",

    # Common
    "sunrise": "sunrise",
    "sunset": "sunset",
}

def get_location_name(lat: float, long: float) -> dict:
    conn = get_db_conn()
    try:
        
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            query = """
                SELECT city_name as name, district, state_ut as region, country
                FROM weatherdata.current_weather_points
                ORDER BY geometry <-> ST_SetSRID(ST_MakePoint(%s, %s), 4326)
                LIMIT 1;
            """

            cursor.execute(query, (long, lat))
            result = cursor.fetchone()

            return dict(result) if result else {}

    except Exception as e:
        return {"error": str(e)}

    finally:
        if conn:
            release_db_conn(conn)
            
def get_icons(condition_text):
    conn = get_db_conn()
    try:
        
        with conn.cursor() as cursor:
            query = """
                SELECT "condition.icon" as icon_url,
                    weatherdata.similarity(
                        lower("condition.text" || ' ' || day_night),
                        lower(%s)
                    ) AS score
                FROM weatherdata.weather_icon_master
                ORDER BY score DESC
                LIMIT 1;
            """

            cursor.execute(query, (condition_text,))
            row = cursor.fetchone()
            icon_url = row[0].replace("//cdn.weatherapi.com/weather/64x64", weather_icon_base_url)
            return icon_url

    except Exception as e:
        return {"error": str(e)}

    finally:
        if conn:
            release_db_conn(conn)

def data_mapping_forecast_weather(data):
    """
    Recursively rename Visual Crossing keys to WeatherAPI keys.
    Works for nested dict/list structures.
    """

    if isinstance(data, list):
        return [data_mapping_forecast_weather(item) for item in data]

    elif isinstance(data, dict):
        new_dict = {}

        for key, value in data.items():

            # Handle condition separately (create nested object)
            if key == "conditions":
                new_dict["condition"] = {
                    "text": value,
                    "icon": get_icons(data.get("icon"))
                }
                continue

            if key == "icon":
                # Skip icon because it's added inside condition
                continue

            new_key = VC_TO_WEATHERAPI_MAP.get(key, key)
            new_dict[new_key] = data_mapping_forecast_weather(value)

        return new_dict

    else:
        return data
 
 
def map_current_vc_to_weatherapi(current):
    return {
        "last_updated": current.get("datetime"),
        "last_updated_epoch": current.get("datetimeEpoch"),

        "temp_c": current.get("temp"),
        "feelslike_c": current.get("feelslike"),
        "humidity": current.get("humidity"),
        "dewpoint_c": current.get("dew"),

        "wind_kph": current.get("windspeed"),
        "wind_degree": current.get("winddir"),
        "wind_dir": None,  # VC doesn't give string like NW

        "pressure_mb": current.get("pressure"),
        "precip_mm": current.get("precip"),
        "vis_km": current.get("visibility"),

        "cloud": current.get("cloudcover"),
        "uv": current.get("uvindex"),

        "gust_kph": current.get("windgust"),

        "condition": {
            "text": current.get("conditions"),
            "icon": get_icons(current.get("icon")) if current.get("icon") else None,
            "code": None
        }
    }
       
def transform_to_weatherapi_format(source):
    if isinstance(source, list):
        forecast_days = source
    elif isinstance(source, dict):
        forecast_days = source.get("forecastday", [])
    else:
        raise ValueError("Invalid source format")

    result = {"forecastday": []}

    for day in forecast_days:
        print("Total_rain_mm:", day.get("precip_mm") ,"Chance_of_rain:", day.get("chance_of_rain"))
        new_day = {
            "date": day.get("date"),
            "date_epoch": day.get("date_epoch"),

            "day": {
                "maxtemp_c": day.get("feelslikemax"),
                "mintemp_c": day.get("feelslikemin"),
                "avgtemp_c": day.get("feelslike_c"),

                "maxwind_kph": day.get("gust_kph"),
                "totalprecip_mm": day.get("precip_mm"),
                "totalsnow_cm": None,

                "avgvis_km": day.get("avgvis_km"),
                "avghumidity": day.get("avghumidity"),

                "daily_chance_of_rain": day.get("chance_of_rain"),

                "condition": {
                    "text": day.get("condition", {}).get("text"),
                    "icon": day.get("condition", {}).get("icon"),
                    "code": None
                },

                "uv": None
            },

            "astro": {
                "sunrise": None,
                "sunset": None,
                "moon_phase": None
            },

            "hour": []
        }

        # Transform hours
        for hr in day.get("hours", []):
            new_hour = {
                "time_epoch": hr.get("date_epoch"),
                "time": f"{day.get('date')} {hr.get('date')}",
                "temp_c": hr.get("temp_c"),
                "is_day": 1 if "day" in hr.get("condition", {}).get("icon", "") else 0,

                "condition": {
                    "text": hr.get("condition", {}).get("text"),
                    "icon": hr.get("condition", {}).get("icon"),
                    "code": None
                },
                
                "wind_kph": hr.get("maxwind_kph"),
                "wind_degree": hr.get("wind_dir"),
                "pressure_mb": hr.get("pressure_mb"),
                "precip_mm": hr.get("precip_mm"),
                "snow_cm": hr.get("totalsnow_cm"),
                "humidity": hr.get("avghumidity"),
                "cloud": hr.get("cloud"),
                "feelslike_c": hr.get("feelslike_c"),
                "dewpoint_c": hr.get("dewpoint_c"),
                "chance_of_rain": hr.get("chance_of_rain"),
                "vis_km": hr.get("avgvis_km"),
                "gust_kph": hr.get("gust_kph"),
                "uv": hr.get("uv")
            }

            new_day["hour"].append(new_hour)

        result["forecastday"].append(new_day)

    return result

def replace_icon_urls_weatherapi(data):
    """
    Recursively traverse JSON and replace 'icon' URLs
    """
    base_url = weather_icon_base_url
    if isinstance(data, dict):
        for key, value in data.items():
            if key == "icon" and isinstance(value, str):
                if value.startswith("//"):
                    value = "https:" + value

                if "/weather/64x64" in value:
                    path = value.split("/weather/64x64")[-1]
                    data[key] = f"{base_url}{path}"
            else:
               replace_icon_urls_weatherapi(value)

    elif isinstance(data, list):
        for item in data:
            replace_icon_urls_weatherapi(item)

    return data

# Forecast data from API
def fetch_forecast_weather_weatherapi(q):
    api_url = "https://api.weatherapi.com/v1/forecast.json"

    params = {
        "key": weatherAPIKey,
        "q": q,
        "days": 8,
        "aqi": "no",
        "alerts": "no"
    }

    response = requests.get(api_url, params=params, timeout=15)
    response.raise_for_status()
    data = response.json()

    # Replace icon URLs
    data = replace_icon_urls_weatherapi(data)
    return data


def fetch_forecast_weather_visualcrossapi(q):
    try:
        base_url = "https://weather.visualcrossing.com/VisualCrossingWebServices/rest/services/timeline"

        location = q
        
        loc = q.split(',')
        latitude = float(loc[0])
        longitude = float(loc[1])

        params = {
            "unitGroup": "metric",
            "include": "current,hours,days",
            "key": visualCrossAPIKey
        }

        url = f"{base_url}/{location}"
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()
        current_data = data.get("currentConditions")
        day_forecast = data.get("days")
        # print("Current Data:", current_data)
        renamed_current_data = map_current_vc_to_weatherapi(current_data)
        # print("Renamed Current Data:", renamed_current_data)
        renamed_day_forecast = data_mapping_forecast_weather(day_forecast)
        location_data = get_location_name(latitude,longitude)
        
        res_data = {
            "location":location_data,
            "current": renamed_current_data,
            "forecast": transform_to_weatherapi_format(renamed_day_forecast)
        }

        return res_data

    except requests.exceptions.RequestException as e:
        return {"error": str(e)}
    
    
# Fetch current weather data
def fetch_weather_data(q, api_name):
    api_name = api_name.lower()

    if api_name == "weatherapi":
        return fetch_forecast_weather_weatherapi(q)
    elif api_name == "visualcrossapi":
        return fetch_forecast_weather_visualcrossapi(q)
    else:
        raise ValueError(f"Unsupported API: {api_name}")
    


