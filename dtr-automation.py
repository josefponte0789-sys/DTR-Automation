import pyautogui
import time
import pandas as pd
import sys
import keyboard



df = pd.read_excel('joborderpersonnels.xlsx')


pyautogui.FAILSAFE = True

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



sys_message = pyautogui.confirm(text='Use the automator?', title='Confirmation', buttons=['OK', 'Cancel'])

if sys_message == 'Cancel':
    exit()

time.sleep(1)

x,y = pyautogui.locateCenterOnScreen('work-timeattendance.PNG',confidence=0.8)
print(x,y)
pyautogui.moveTo(x,y)
pyautogui.click()


time.sleep(1.8)

x,y = pyautogui.locateCenterOnScreen('schoolname.PNG',confidence=0.8)
print(x,y)
pyautogui.moveTo(x,y)
pyautogui.click()


time.sleep(0.5)


pyautogui.write('10000')

time.sleep(1.5)

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
date_picker = pyautogui.prompt(text='Pick a Date(YYYY-MM-DD)', title='Date Picker', default='')

selection_list = pyautogui.prompt(text='Type "jo" for Job Order Personnel, "nurse" for Nurse personnel.',title='What are we printing?',default='')


if(selection_list == 'jo'):
    for personnel_names in df['JobOrderName']:
        if keyboard.is_pressed('q'):
            sys.exit()
        
        time.sleep(0.05)
        
        if pd.isna(personnel_names):
            break
        
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

        time.sleep(0.5)
        
        print_ic = pyautogui.locateCenterOnScreen('grid-print.PNG', confidence=0.8)
        print(print_ic)
        pyautogui.moveTo(print_ic)
        pyautogui.click(x=1120,y=260)

        time.sleep(0.4)

        pyautogui.hotkey('ctrl','a')

        time.sleep(0.08)
        pyautogui.press('backspace')
        time.sleep(0.08)
        pyautogui.write(date_picker)

        time.sleep(0.7)
        
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


elif(selection_list == 'nurse'):
         for personnel_names in df['NurseName']:
            if keyboard.is_pressed('q'):
                sys.exit()
            
            time.sleep(0.05)
             
            if pd.isna(personnel_names):
                break
                
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

            time.sleep(0.5)
            
            print_ic = pyautogui.locateCenterOnScreen('grid-print.PNG', confidence=0.8)
            print(print_ic)
            pyautogui.moveTo(print_ic)
            pyautogui.click(x=1120,y=260)

            time.sleep(0.4)

            pyautogui.hotkey('ctrl','a')

            time.sleep(0.08)
            pyautogui.press('backspace')
            time.sleep(0.08)
            pyautogui.write(date_picker)

            time.sleep(0.7)
            
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
            
if selection_list == 'Cancel':
    exit()


        

   