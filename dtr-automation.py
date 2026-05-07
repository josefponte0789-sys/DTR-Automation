import pyautogui
import time
import pandas as pd


df = pd.read_excel('joborderpersonnels.xlsx')






sys_message = pyautogui.confirm(text='Use the automator?', title='Confirmation', buttons=['OK', 'Cancel'])

if sys_message == 'Cancel':
    exit()

time.sleep(1)

#########################################

##TESTING CHANGING OF DATES##
# print_ic = pyautogui.locateCenterOnScreen('printicon.PNG', confidence=0.8)
# print(print_ic)
# if print_ic is not None:
#     pyautogui.moveTo(print_ic)
#     pyautogui.click()
# else:
#     print('Skipping personnel...')

# time.sleep(1)
# print_ic = pyautogui.locateCenterOnScreen('view-print.PNG', confidence=0.8)
# print(print_ic)
# pyautogui.moveTo(print_ic)
# pyautogui.click()

########################################



x,y = pyautogui.locateCenterOnScreen('schoolname.PNG',confidence=0.8)
print(x,y)
pyautogui.moveTo(x,y)
pyautogui.click()


time.sleep(0.5)


pyautogui.write('10000')

time.sleep(2.2)
schoolid_name = pyautogui.locateCenterOnScreen('schoolid.PNG', confidence=0.8)
print(schoolid_name)
pyautogui.moveTo(schoolid_name)
pyautogui.click()


time.sleep(6.5)

search_loc = pyautogui.locateCenterOnScreen('searchbar.PNG', confidence=0.8)
print(search_loc)
pyautogui.moveTo(search_loc)
pyautogui.click()


time.sleep(1)


for personnel_names in df['Name']:
    
    print(f'Processing: {personnel_names}')
    pyautogui.write(personnel_names)
    pyautogui.press('Enter')

    print_ic = pyautogui.locateCenterOnScreen('printicon.PNG', confidence=0.8)
    print(print_ic)
    if print_ic is not None:
        pyautogui.moveTo(print_ic)
        pyautogui.click()
    else:
        print('Skipping personnel...')
        continue

    time.sleep(1)

    print_ic = pyautogui.locateCenterOnScreen('view-print.PNG', confidence=0.8)
    print(print_ic)
    pyautogui.moveTo(print_ic)
    pyautogui.click()


    time.sleep(3)

    download_df = pyautogui.locateCenterOnScreen('downloadpdf.PNG', confidence=0.8)
    print(download_df)
    pyautogui.moveTo(download_df)
    pyautogui.click()
    time.sleep(1.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','w')


    time.sleep(0.5)

    close_timesheet = pyautogui.locateCenterOnScreen('close-printtimesheet.PNG', confidence=0.8)
    print(close_timesheet)
    pyautogui.moveTo(close_timesheet)
    pyautogui.click()
    
    time.sleep(1)
    search_loc = pyautogui.locateCenterOnScreen('searchbar.PNG', confidence=0.8)
    print(search_loc)
    pyautogui.moveTo(search_loc)
    pyautogui.click()
    pyautogui.hotkey('ctrl','a')
    time.sleep(0.5)
    pyautogui.press('backspace')
    time.sleep(0.5)



