import logging
import os
import json
import pandas as pd
import requests
from datetime import datetime
from logging.handlers import RotatingFileHandler
from replicated_script import user_input  



# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(funcName)s(%(lineno)d) %(message)s',
    handlers=[
        RotatingFileHandler("logs.txt", maxBytes=5*1024*1024, backupCount=2),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

api_key = os.environ.get("API_KEY_MAPS")

directory = os.getcwd()

user_name = user_input['username']
origin = user_input['origin']
destination = user_input['destination'] 


def get_route_info(origin, destination, api_key):
    endpoint = 'https://maps.googleapis.com/maps/api/directions/json?'
    nav_request = f'origin={origin}&destination={destination}&key={api_key}'
    response = requests.get(endpoint + nav_request)
    directions = json.loads(response.text)
    return directions

def extract_travel_time(route_info):
    if not route_info or 'routes' not in route_info or not route_info['routes']:
        logger.error("No routes found in the response.")
        return None, None

    first_route = route_info['routes'][0]
    if 'legs' not in first_route or not first_route['legs']:
        logger.error("No legs found in the route.")
        return None, None

    legs = first_route['legs']
    total_duration = sum(leg['duration']['value'] for leg in legs if 'duration' in leg)
    total_duration_text = " / ".join(leg['duration']['text'] for leg in legs if 'duration' in leg)
    return total_duration, total_duration_text

def cleaned_origin_and_destination(origin, destination):
    cleaned_origin = origin.replace(" ", "_")
    cleaned_destination = destination.replace(" ", "_") 
    return cleaned_origin, cleaned_destination

def save_to_excel(data, base_dir, cleaned_origin, cleaned_destination):
    file_name = f"{cleaned_origin}_to_{cleaned_destination}.xlsx"
    file_path = os.path.join(base_dir, file_name)
    df_new = pd.DataFrame([data])
    if os.path.exists(file_path):
        df_existing = pd.read_excel(file_path)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new
    df_combined.to_excel(file_path, index=False)
    logger.info(f"Data saved to {file_path}")
    return file_path

def run_script(origin, destination):
    base_dir = os.getcwd()
    cleaned_origin, cleaned_destination = cleaned_origin_and_destination(origin, destination)

    if not os.path.exists(base_dir):
        logger.info("Folder not found")

    if api_key is not None:
        try:
            route_info = get_route_info(origin, destination, api_key)
            total_duration, total_duration_text = extract_travel_time(route_info)
            if total_duration is not None:
                   data_to_save = {
                    'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'origin': origin,
                    'destination': destination,
                    'Total Duration (sec)': total_duration,
                    'Total Duration (text)': total_duration_text
                    }
                   save_to_excel(data_to_save, base_dir, cleaned_origin, cleaned_destination)                  
                            
            else:
                    data_to_save = {
                        'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'Error': 'API call succeeded but no route found'
                    }
                    save_to_excel(data_to_save, base_dir, cleaned_origin, cleaned_destination)

        except Exception as e:
                logger.error("API call failed: " + str(e))
                data_to_save = {
                    'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Error': 'API call failed'
                }
                save_to_excel(data_to_save, base_dir, origin, destination)
if __name__ == '__main__':
    run_script(origin, destination)
