import pyautogui
import time
import pandas as pd
import sys
import keyboard
import threading
from tkinter import *



df = pd.read_excel('joborderpersonnels.xlsx')
running = True


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


def start_automation():
    global running
    sys_message = pyautogui.confirm(text='Use the automator?', title='Confirmation', buttons=['OK', 'Cancel'])

    if sys_message == 'Cancel':
        exit()
        

    time.sleep(1)
    while running:
        try:
            result_label.config(text=f"")

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


            time.sleep(7.5)

            search_loc = pyautogui.locateCenterOnScreen('searchbar.PNG', confidence=0.8)
            print(search_loc)
            pyautogui.moveTo(search_loc)
            pyautogui.click()


            time.sleep(1)
            date_picker = pyautogui.prompt(text='Pick a Date(YYYY-MM-DD)', title='Date Picker', default='')

            selection_list = pyautogui.prompt(text='Type "jo" for Job Order Personnel, "nurse" for Nurse personnel.',title='What are we saving?',default='')
            start_from = pyautogui.prompt(text='Start saving from which personnel?', title='Start saving from?', default='')
            start_indexjo = df.index[df['JobOrderName'] == start_from][0]
            subset = df.iloc[start_indexjo:]



            if(selection_list == 'jo'):
                for personnel_names in subset['JobOrderName']:
                    if not running:
                        break
                    
                    if keyboard.is_pressed('q'):
                        sys.exit()
                    
                    time.sleep(0.75)
                    
                    if pd.isna(personnel_names):
                        time.sleep(1.5)
                        result_label.config(text=f"Finished.")
                        break
                    
                    main_label.config(text=f'Click the "quit" button to exit the program!')
                    
                    
                    print(f'Processing: {personnel_names}')
                    result_label.config(text=f"Saving: {personnel_names}", font=('Comfortaa', 12, 'bold'))
                    root.update()
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

                    time.sleep(1.5)
                    
                    print_ic = pyautogui.locateCenterOnScreen('view-print.PNG', confidence=0.8)
                    print(print_ic)
                    pyautogui.moveTo(print_ic)
                    pyautogui.click()


                    time.sleep(3.5)

                    download_df = pyautogui.locateCenterOnScreen('downloadpdf.PNG', confidence=0.8)
                    print(download_df)
                    time.sleep(0.75)
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
                        
                        time.sleep(0.75)
                        
                        if pd.isna(personnel_names):
                            time.sleep(1.5)
                            result_label.config(text=f"Finished.")
                            break
                        
                        main_label.config(text=f'Click the "quit" button to exit the program!')
                        
                        print(f'Processing: {personnel_names}')
                        result_label.config(text=f"Saving: {personnel_names}", font=('Comfortaa', 12, 'bold'))
                        root.update()
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

                        time.sleep(1.5)
                        
                        print_ic = pyautogui.locateCenterOnScreen('view-print.PNG', confidence=0.8)
                        print(print_ic)
                        pyautogui.moveTo(print_ic)
                        pyautogui.click()


                        time.sleep(3.5)

                        download_df = pyautogui.locateCenterOnScreen('downloadpdf.PNG', confidence=0.8)
                        print(download_df)
                        time.sleep(0.75)
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
        except pyautogui.ImageNotFoundException:
            result_label.config(text=f'Image Error has occured! Restarting...')
            time.sleep(3)
            
        except Exception:
            result_label.config(text=f'An error has occured. Breaking...')
            time.sleep(3)
            running = False


            
def start_script():
    global running
    running = True
    threading.Thread(target = start_automation, daemon=True).start()


def stop_script():
    global running
    running = False
    print("Stopped!")
    
    
root = Tk()
root.geometry('250x280')

root.attributes('-topmost', True)


main_label = Label(root, text='')
main_label.pack(pady=10)

start_btn = Button(root, text='Start Automation', command=start_script)
start_btn.pack(pady=10)

stop_btn = Button(root, text='Stop Automation', command=stop_script)
stop_btn.pack(pady=10)

quit_btn = Button(root, text='Quit', command=root.destroy)
quit_btn.pack(pady=10)



result_label = Label(root, text='')
result_label.pack(pady=10)


root.mainloop()