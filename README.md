# Visit MSN Weather

This script automates the task of visiting the MSN Weather website at specific times during the day and records whether the task was completed or not. It also allows the user to toggle whether the script starts automatically at startup using the system tray icon.

## Requirements
- Selenium
- pystray
- PIL
- winreg

## Usage
1. Install the required packages by running `pip install -r requirements.txt`
2. Run the script `python visit_msn_weather.py`
3. The script will run in the background and visit the MSN Weather website at the specified times. You can stop the script by right-clicking on the system tray icon and selecting "Quit"
4. You can also toggle whether the script starts automatically at startup by right-clicking on the system tray icon and selecting "Startup"

## Note
* The script uses the time and datetime libraries to check the current time and decide whether to run the task or not.
* The script uses the random library to generate random sleep time.
* The script uses the threading library to run the scheduler as a separate thread.
* The script uses the winreg library to check whether the script is already set to run at startup and also toggle the script to run at startup.
* The script uses the Selenium library to automate the process of visiting the website.
* The script uses the pystray library to create an icon in the system tray that can be used to stop the script or toggle whether the script starts automatically at startup.


