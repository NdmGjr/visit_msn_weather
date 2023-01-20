import os
import ast
import time
import random
import winreg
import threading
import pystray
from PIL import Image
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

app_name = "Visit MSN Weather"
sleep = random.uniform(3, 4)

def visit_msn_weather(task=0, close=True):
    options = webdriver.ChromeOptions()
    options.add_argument("--app=https://www.msn.com/en-us/weather/forecast/in-Punjab,Pakistan")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("--window-size=1366,768")
    options.add_argument("--window-position=-1366,0")
    # options.add_argument("--user-data-dir=C:/Users/ndmgo/AppData/Local/Google/Chrome/User Data")
    # options.add_argument("--profile-directory=Profile 8") #user profile option ain't working for some reason
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 60)
    try:
        # Navigate to login page
        driver.get("https://www.msn.com/en-us/weather/forecast/in-Punjab,Pakistan")
        try:
            login_button = wait.until(EC.visibility_of_element_located((By.ID, "mectrl_main_trigger")))
            time.sleep(sleep)
            login_button.click()
            email_input = wait.until(EC.visibility_of_element_located((By.NAME, "loginfmt")))
            email_input.send_keys("ndmgorsi@outlook.com")
            next_button = wait.until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
            time.sleep(sleep)
            next_button.click()
            password_input = wait.until(EC.visibility_of_element_located((By.NAME, "passwd")))
            password_input.send_keys("microsoft_account_password")

            sign_in_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            time.sleep(sleep)
            sign_in_button.click()
            try:
                stay_signed_in_div = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(),'Stay signed in')]")))
                checkbox = driver.find_element(By.ID, "KmsiCheckboxField")
                if checkbox.is_enabled():
                    time.sleep(sleep)
                    checkbox.click()
                    sign_in_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
                    time.sleep(sleep)
                    sign_in_button.click()
            except TimeoutException:
                pass
        except TimeoutException:
            print("Already Logged in")
        time.sleep(5)

        if task == 0:
            time.sleep(5)
            try_to_click(wait, By.XPATH, "//*[@id='pivotsNav']/div/div/a[1]")
            # try_to_click(wait, By.XPATH, "//*[@id='WeatherLifeIndexEntry-ScreenWidth-c4']/div/div/div[2]/div[2]/div/a[1]")
            try_to_click(wait, By.XPATH, "//ul[@class='cardContainer-E1_2']/li[2]/button")
            try_to_click(wait, By.XPATH, "//button[@data-value='hourly']")
            try_to_click(wait, By.XPATH, "//*[@id='ForecastHourly']/div/ul/li[1]/button")
            try_to_click(wait, By.XPATH, "//*[@id='pivotsNav']/div/div/a[2]")
            time.sleep(5)
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-0-g_temp']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-1-g_precip']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-2-g_radar']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-3-g_wind']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-4-g_cloud']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-5-g_humidity']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-6-g_visibility']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-7-g_pressure']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-8-g_dewpoint']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-9-g_air_quality']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-0-g_hurricane']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-1-g_wildfire']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-2-g_winter_storm']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-3-g_severe_wea']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-4-g_earthquake']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-0-g_ski_resort']")
            try_to_click(wait, By.XPATH, "//*[@id='layer-group-0-g_etree']")
            
            try_to_click(wait, By.XPATH, "//*[@id='pivotsNav']/div/div/a[4]")
            try_to_click(wait, By.XPATH, "//*[@id='pivotsNav']/div/div/a[5]")
            try_to_click(wait, By.XPATH, "//*[@id='WeatherMonthCalendarSection']/div/div[2]/div/ul/li[2]/a")
            time.sleep(5)
            try_to_click(wait, By.XPATH, "//*[@id='pivotsNav']/div/div/a[8]")

            # try_to_click(wait, By.XPATH, "//*[@id='overflowBtnId']/button")
            # try_to_click(wait, By.XPATH, "//*[@id='pivotsNav']/div/div/a[12]")
        elif task == 1:
            # try_to_click(wait, By.XPATH, "//*[@id='root']/div/div/div[3]/div[0]/div[5]/div[1]/div/div[1]/button")
            close = False
        elif task == 2:
            # try_to_click(wait, By.XPATH, "//*[@id='root']/div/div/div[3]/div[1]/div[5]/div[1]/div/div[1]/button")
            close = False
        elif task == 3:
            # try_to_click(wait, By.XPATH, "//*[@id='root']/div/div/div[3]/div[2]/div[5]/div[1]/div/div[1]/button")
            close = False

        if not close:
            current_handles = driver.window_handles
            WebDriverWait(driver, 60*60).until(lambda d: current_handles != d.window_handles) # wait for 1 hour
        driver.quit()
        return True
    except Exception as e:
        print(f"An error occurred: {e}")
        driver.quit()

def try_to_click(wait, by, value):
    try:
        time.sleep(sleep)
        element = wait.until(EC.element_to_be_clickable((by, value)))
        time.sleep(sleep)
        element.click()
    except Exception as e:
        print(f"An error occurred: {e}")
        print(f"Element with {by} = {value} didn't work")


def get_all_tasks():
    with open("E:/NdmGjr/visit_msn_weather/vars.txt", "r") as file:
        data = file.read().split("\n")
        stored_date = datetime.strptime(data[4], "%Y-%m-%d").date()
        current_date = datetime.now().date()
        tasks = data[:4]
        if current_date > stored_date:
            tasks = ["False" for i in range(4)]
            with open("E:/NdmGjr/visit_msn_weather/vars.txt", "w") as file:
                file.write("\n".join(tasks + [current_date.strftime("%Y-%m-%d")]))
    return tasks

def update_vars(task):
    tasks = get_all_tasks()
    tasks[task] = "True"
    current_date = datetime.now().date()
    with open("E:/NdmGjr/visit_msn_weather/vars.txt", "w") as file:
        file.write("\n".join(tasks + [current_date.strftime("%Y-%m-%d")]))

def right_time(task):
    current_hour = time.localtime().tm_hour
    if task == 0:
        return True
    elif task == 1:
        return 8 <= current_hour < 10
    elif task == 2:
        return 12 <= current_hour < 14
    elif task == 3:
        return 16 <= current_hour < 18
    else:
        return False

def run_scheduler(stop_event):
    while not stop_event.is_set():
        tasks = [ast.literal_eval(x) for x in get_all_tasks()]
        for i, task in enumerate(tasks):
            if not task and right_time(i):
                if visit_msn_weather(i):
                    update_vars(i)
        time.sleep(60*30)

# start the scheduler thread
stop_event = threading.Event()
scheduler_thread = threading.Thread(target=run_scheduler, args=(stop_event,))
scheduler_thread.start()

#icon related code

def on_quit():
    stop_event.set()
    scheduler_thread.join()
    icon.stop()

def get_startup_status():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_ALL_ACCESS)
    try:
        value = winreg.QueryValueEx(key, app_name)[0]
        return True
    except:
        return False

def startup_toggle(icon, item):
    global startup_status
    startup_status = not item.checked
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_ALL_ACCESS)
    if startup_status:
        value = os.path.join(os.getcwd(), "visit_msn_weather.exe")
        winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, value)
    else:
        winreg.DeleteValue(key, app_name)

startup_status = get_startup_status()

image = Image.open('E:/NdmGjr/visit_msn_weather/icon.png')
icon = pystray.Icon(app_name, image, app_name)
icon.menu = pystray.Menu(
    pystray.MenuItem(text="MSN Weather", action=lambda: visit_msn_weather(None, False), default=True),
    pystray.MenuItem("Quit", on_quit),
    pystray.MenuItem("Run at startup", startup_toggle, checked=lambda item: startup_status)
)

icon.run()

# pyinstaller --name visit_msn_weather main.py --icon=icon.png --noconsole --onefile

