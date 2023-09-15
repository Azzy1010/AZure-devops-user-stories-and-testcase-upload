from datetime import datetime, date
import time
import pyautogui


driver = pyautogui
try:
    driver.hotkey('win')
    time.sleep(.5)
    driver.write('Slack')
    driver.press('enter')
    time.sleep(1)
    driver.hotkey('ctrl', 'g')
    time.sleep(0.5)
    try:
        driver.hotkey('ctrl', 'a')
        driver.press('backspace')
        time.sleep(0.3)
    except:
        pass
    driver.write('Azfer Muhammad Gul')
    time.sleep(1)
    for i in range(3):
        driver.press('down')
    time.sleep(.3)
    driver.press('enter')
    now = datetime.now()
    current_time1 = now.strftime("%H:%M")
    today = date.today().strftime("%d/%m/%Y")
    if int(current_time1.split(':')[0]) > 11:
        current_time = now.strftime("%I:%M")
        driver.write(f'{today} --> Time out: {current_time}')
    else:
        current_time = now.strftime("%I:%M")
        driver.write(f'{today} --> Time in: {current_time}')
    time.sleep(0.5)
    driver.press('enter')
    time.sleep(2)
    driver.hotkey('alt', 'f4')
    driver.alert(text='Time sent Successfully!', title='Alert!')
    print('Time sent Successfully!')
except Exception as e:
    print(f'Failed to send time {e}')
