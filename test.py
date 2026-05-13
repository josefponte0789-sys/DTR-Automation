import pyautogui
import time
import pandas as pd
from tkinter import *
from tkinter import ttk
import keyboard
import sys
import threading


running = True

df = pd.read_excel('joborderpersonnels.xlsx')


# #### MOUSE COORDINATOR LOCATOR TOOL ####
# import pyautogui, sys
# print('Press Ctrl-C to quit.')
# try:
#     while True:
#         x, y = pyautogui.position()
#         positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
#         print(positionStr, end='')
#         print('\b' * len(positionStr), end='', flush=True)
# except KeyboardInterrupt:
#     print('\n')

def show_text():
    global running
    
    for personnel_names in df['NurseName']:
            if not running:
                result_label.config(text=f"Stopped.")
                print('stopped')
                break
            if keyboard.is_pressed('q'):
                sys.exit()
            
            time.sleep(0.05)
            
            if pd.isna(personnel_names):
                time.sleep(1.5)
                result_label.config(text=f"Finished.")
                break
                
            
     
            
            
            print(f'Processing: {personnel_names}')
            result_label.config(text=f"Saving: {personnel_names}")
            root.update()
            pyautogui.write(personnel_names)
            pyautogui.press('Enter')


def start_script():
    global running
    running = True
    threading.Thread(target=show_text,daemon=True).start()
  
def stop_script():
    global running
    running = False
            
root = Tk()
root.geometry('320x320')

btn = Button(root,text='click me!',command = start_script)

btn.pack(pady=10)


stop_btn = Button(root, text='quit', command=stop_script)
stop_btn.pack(pady=10)

result_label = Label(root, text='')
result_label.pack(pady=10)



root.mainloop()








