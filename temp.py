#
# user_url = pyautogui.prompt(text= 'Please provide URl to open: ')
# def browser_open():
#     pyautogui.hotkey('win')
#     time.sleep(.5)
#     pyautogui.write('cmd')
#     pyautogui.press('enter')
#     time.sleep(1)
#     pyautogui.write('cd "C:\Program Files\Google\Chrome\Application"')
#     pyautogui.press('enter')
#     pyautogui.write('chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Chromedata"')
#     pyautogui.press('enter')
#     pyautogui.write('exit')
#     pyautogui.press('enter')
#
#
# def chrome_config():
#     global driver
#     browser_open()
#     options = Options()
#     options.add_experimental_option("debuggerAddress", "localhost:9222")
#     driver = webdriver.Chrome(options=options)
#     driver.maximize_window()
#
# chrome_config()
# driver.get(user_url)

# from pathlib import Path
# print(type(str(Path.cwd())))
# ss_directory = f"./Reports/hi/SS"
# os.makedirs(ss_directory, exist_ok=True)
# print(Path.cwd())