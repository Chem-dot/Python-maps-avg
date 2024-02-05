from flask import Flask, request, render_template
import pandas as pd
import os
import requests
import json
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler
import shutil
import win32com.client
import pythoncom
from datetime import timedelta

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

directory = os.getcwd()
app = Flask(__name__,static_folder='static')

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

def store_inputs(base_dir, user_inputs):
    file_path = os.path.join(base_dir, 'replicated_script.py')
    with open(file_path, 'w') as file:
        file.write(f"user_input = {user_inputs}\n")
    logger.info(f"Inputs stored in {file_path}: {user_inputs}")

def scriptcopy(cleaned_origin, cleaned_destination, base_dir):
    script_name = "templates/pythoncopy.py"
    copied_script_name = f"{cleaned_origin}_to_{cleaned_destination}.py"
    try:
        shutil.copy(script_name, os.path.join(base_dir, copied_script_name))
        logger.info(f"{script_name} has been copied to {base_dir} with the name {copied_script_name}")
        return copied_script_name
    except Exception as e:
        logger.error("Script duplication failed: " + str(e))
        data_to_save = {
            'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Error': 'Script duplication failed'
        }
        save_to_excel(data_to_save, base_dir, cleaned_origin, cleaned_destination)

def delete_inputs(base_dir):
    file_path = os.path.join(base_dir, 'replicated_script.py')
    try:
        os.remove(file_path)
        logger.info(f"File {file_path} has been removed")
    except Exception as e:
        logger.error("File removal failed: " + str(e))

def schedule_task(base_dir, cleaned_origin, cleaned_destination, copied_script_name):
    try:
        pythoncom.CoInitialize()
        # Initialize the constants
        win32com.client.gencache.EnsureDispatch('Schedule.Service')
        TASK_TRIGGER_DAILY = 2
        TASK_CREATE_OR_UPDATE = 6
        TASK_ACTION_EXEC = 0

        scheduler = win32com.client.Dispatch('Schedule.Service')
        scheduler.Connect()

        root_folder = scheduler.GetFolder('\\')

        task_def = scheduler.NewTask(0)

        start_time = datetime.now() + timedelta(minutes=1)
        trigger = task_def.Triggers.Create(TASK_TRIGGER_DAILY)
        trigger.StartBoundary = start_time.isoformat()

        action = task_def.Actions.Create()

        # Check if the action is an ExecAction
        if action.Type == win32com.client.constants.TASK_ACTION_EXEC:
            # Now we can safely set the Path and Arguments properties
            action.Path = "C:\\Users\\info\\AppData\\Local\\Microsoft\\WindowsApps\\python.exe"
            action.Arguments = f'{base_dir}, {copied_script_name}'
            return action.path, action.arguments
        else:
            logger.error(f"Unexpected action type: {action.Type}")

        action = task_def.Actions.Create(TASK_ACTION_EXEC)
        action.Path = "C:\\Users\\info\\AppData\\Local\\Microsoft\\WindowsApps\\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\\python.exe"
        action.Arguments = f'"{copied_script_name}"'
        logger.info(action.arguments)

        task_def.RegistrationInfo.Description = f"Task to run {copied_script_name} at {base_dir}"
        task_def.Settings.Enabled = True
        task_def.Settings.Hidden = False
        task_def.Settings.StopIfGoingOnBatteries = False        

        # Set the logon type and user
        principal = task_def.Principal
        principal.UserId = "info@chonathanit.com",  # User
        principal.password = "@EV8FXaZ6YBUeWwrEUWr",  # Password
        principal.LogonType = 1  # TASK_LOGON_PASSWORD

        root_folder.RegisterTaskDefinition(
            f"{cleaned_origin}_to_{cleaned_destination}",  # Task name
            task_def,
            TASK_CREATE_OR_UPDATE,
            "info@chonathanit.com",  # User
            "@EV8FXaZ6YBUeWwrEUWr",  # Password
            win32com.client.constants.TASK_LOGON_PASSWORD,  # Logon type
            )

        logger.info(f'Task scheduled to run {base_dir} at {start_time}')
    except Exception as e:
        logger.error("Task scheduling failed: " + str(e))
        data_to_save = {
        'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Error': 'Task scheduling failed ' + str(e)
    }
    save_to_excel(data_to_save, base_dir, cleaned_origin, cleaned_destination)
    

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
    base_dir = os.path.join(directory,'users', user_name)
    cleaned_origin, cleaned_destination = cleaned_origin_and_destination(origin, destination)
    copied_script_name = scriptcopy(cleaned_origin, cleaned_destination, base_dir)

    user_inputs = {
        'username' : user_name,
        'origin': origin,
        'destination': destination
    }

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
                   save_to_excel(data_to_save, base_dir, cleaned_origin, cleaned_destination)
                   store_inputs(base_dir, user_inputs)
                   scriptcopy(cleaned_origin, cleaned_destination, base_dir)
                   schedule_task(base_dir, cleaned_origin, cleaned_destination, copied_script_name)
                   #delete_inputs(base_dir)                  
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
        return render_template('return.html')

    if 'average_button' in request.form:
        if os.path.exists(base_dir):
            excel_file_path = os.path.join(base_dir, f"{cleaned_origin}_to_{cleaned_destination}.xlsx")
            df = pd.read_excel(excel_file_path)
            avg_e = int(df['Total Duration (sec)'].iloc[1:].mean() / 60)
            Origin = df.at[1, 'origin'].title()
            Destination = df.at[1, 'destination'].title()
            num_rows = df.shape[0]
            logger.info(f"The average travel time from {Origin} to {Destination} is: {avg_e} minutes over the past {num_rows} days")
            print(f"The average travel time from {Origin} to {Destination} is: {avg_e} minutes over the past {num_rows} days")
        return render_template('average.html', Origin=Origin, Destination=Destination, avg_e=avg_e, num_rows=num_rows)
    else:
        logger.info(f"The path {excel_file_path} had an issue")
if __name__ == '__main__':
    app.run(debug=False)

   