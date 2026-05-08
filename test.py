import pyautogui
import time
import pandas as pd
from tkinter import *
from tkinter import ttk
import keyboard
import sys



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
    for personnel_names in df['NurseName']:
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
  
            
root = Tk()
root.geometry('320x320')

btn = Button(root,text='click me!',command = show_text)
btn.pack(pady=10)

result_label = Label(root, text='')
result_label.pack(pady=10)



root.mainloop()








