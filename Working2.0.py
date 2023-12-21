from flask import Flask, request, render_template
import pandas as pd
import os
import requests
import json
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler
import sys

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

app = Flask(__name__)

api_key = os.environ.get("API_KEY_MAPS")

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

def save_to_excel(data, base_dir, origin, destination):
    formatted_origin = origin.replace(" ", "_")
    formatted_destination = destination.replace(" ", "_")
    file_name = f"{formatted_origin}_to_{formatted_destination}.xlsx"
    file_path = os.path.join(base_dir, file_name)

    df_new = pd.DataFrame([data])

    if os.path.exists(file_path):
        df_existing = pd.read_excel(file_path)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new

    df_combined.to_excel(file_path, index=False)
    return file_path

@app.route('/')
def index():
    return render_template('home.html')

@app.route('/home.html')
def home():
    return render_template('home.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = request.form
    user_name = data.get('name').strip()
    origin = data.get('origin')
    destination = data.get('destination')
    base_dir = os.path.join('C:/Users/info/OneDrive/Documents/Code stuff/users', user_name)

    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
        logger.info(f"Folder created for user: {user_name}")

    if 'submit_button' in request.form and api_key is not None:
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
                   saved_file_path = save_to_excel(data_to_save, base_dir, origin, destination)
                   logger.info(f"Data saved to {saved_file_path}")
            else:
                    data_to_save = {
                        'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'Error': 'API call succeeded but no route found'
                    }
                    saved_file_path = save_to_excel(data_to_save, base_dir, origin, destination)
                    logger.info(f"Data saved to {saved_file_path}")
        except Exception as e:
                logger.error("API call failed: " + str(e))
                data_to_save = {
                    'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Error': 'API call failed'
                }
                saved_file_path = save_to_excel(data_to_save, base_dir, origin, destination)
                logger.info(f"Data saved to {saved_file_path}")
        return render_template('return.html')

    if 'average_button' in request.form:
        if os.path.exists(base_dir):
            excel_file_path = os.path.join(base_dir, f"{origin}_to_{destination}.xlsx")
            df = pd.read_excel(excel_file_path)
            avg_e = int(df['Total Duration (sec)'].iloc[1:].mean() / 60)
            Origin = df.at[1, 'origin']
            Destination = df.at[1, 'destination']
            num_rows = df.shape[0]
            logger.info(f"The average travel time from {Origin} to {Destination} is: {avg_e} minutes over the past {num_rows} days")
            print(f"The average travel time from {Origin} to {Destination} is: {avg_e} minutes over the past {num_rows} days")
        return render_template('average.html', Origin=Origin, Destination=Destination, avg_e=avg_e, num_rows=num_rows)

if __name__ == '__main__':
    app.run(debug=True)
