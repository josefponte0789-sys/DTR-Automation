import pyautogui
import time
import pandas as pd
import sys
import keyboard
import threading
from tkinter import *

# ── Config ────────────────────────────────────────────────────────────────────

EXCEL_FILE = 'joborderpersonnels.xlsx'
PERSONNEL_TYPES = {
    'jo':    'JobOrderName',
    'nurse': 'NurseName',
}

pyautogui.FAILSAFE = True

# ── State ─────────────────────────────────────────────────────────────────────

df = pd.read_excel(EXCEL_FILE)
running = False

# ── Helpers ───────────────────────────────────────────────────────────────────

def pick_date():
    return pyautogui.prompt(text='Pick a Date (YYYY-MM-DD)', title='Date Picker', default='')


def find_and_click(image, confidence=0.8, retries=1):
    """Locate an image on screen and click it. Returns the center point or None."""
    for _ in range(retries):
        loc = pyautogui.locateCenterOnScreen(image, confidence=confidence)
        if loc:
            pyautogui.moveTo(loc)
            pyautogui.click()
            return loc
    return None


def navigate_to_search():
    """Navigate to the time attendance search page and return when the search bar is ready."""
    find_and_click('work-timeattendance.PNG', confidence=0.7)
    time.sleep(2.5)

    find_and_click('schoolname.PNG', confidence=0.8)
    time.sleep(1.5)

    pyautogui.write('10000')
    time.sleep(2.5)
    pyautogui.press('Enter')
    time.sleep(10.5)

    find_and_click('searchbar.PNG', confidence=0.8)
    time.sleep(1)


def clear_search():
    """Clear the search bar and return focus to it."""
    find_and_click('searchbar.PNG', confidence=0.8)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('backspace')
    time.sleep(0.5)


def process_one(name, date_picker):
    """
    Run the full print-and-download flow for a single personnel name.
    Returns True on success, False if the personnel should be skipped.
    """
    pyautogui.write(name)
    pyautogui.press('Enter')

    # Click the print icon — skip if not found
    if not find_and_click('printicon.PNG', confidence=0.8):
        print(f'  No print icon found — skipping {name}')
        return False

    time.sleep(0.5)

    # Click the grid-print button with a small offset
    grid_loc = pyautogui.locateCenterOnScreen('grid-print.PNG', confidence=0.8)
    if grid_loc is None:
        return False
    time.sleep(0.7)
    pyautogui.click(grid_loc.x - 60, grid_loc.y - 10)
    time.sleep(0.4)

    # Enter the chosen date
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.08)
    pyautogui.press('backspace')
    time.sleep(0.08)
    pyautogui.write(date_picker)
    time.sleep(1.5)

    # View / print
    find_and_click('view-print.PNG', confidence=0.9)
    time.sleep(3.5)

    # Download
    find_and_click('downloadpdf.PNG', confidence=0.8)
    time.sleep(1.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(0.5)

    # Close the timesheet panel
    find_and_click('close-printtimesheet.PNG', confidence=0.8)
    time.sleep(1)

    return True


# ── Core automation ───────────────────────────────────────────────────────────

def run_automation():
    """Main automation loop — handles navigation, personnel iteration, and restarts."""
    global running

    personnel_type = pyautogui.prompt(
        text='Type "jo" for Job Order Personnel, "nurse" for Nurse personnel.',
        title='What are we saving?',
        default=''
    )
    if not personnel_type or personnel_type not in PERSONNEL_TYPES:
        set_status('Invalid personnel type. Stopping.')
        running = False
        return

    date_picker = pick_date()
    if not date_picker:
        set_status('No date selected. Stopping.')
        running = False
        return

    col = PERSONNEL_TYPES[personnel_type]

    while running:
        try:
            navigate_to_search()

            start_from = pyautogui.prompt(
                text='Start saving from which personnel?',
                title='Start saving from?',
                default=''
            )
            start_idx = df.index[df[col] == start_from][0]
            subset = df.iloc[start_idx:]

            for name in subset[col]:
                if not running or keyboard.is_pressed('q'):
                    sys.exit()

                time.sleep(0.75)

                if pd.isna(name):
                    set_status('Finished.')
                    running = False
                    break

                print(f'Processing: {name}')
                set_status(f'Saving: {name}')

                process_one(name, date_picker)
                clear_search()

            break  # Clean exit after finishing the list

        except pyautogui.ImageNotFoundException:
            set_status(f'Image not found — restarting from last checkpoint...')
            time.sleep(3)
            # Loop continues: will re-navigate and prompt for start_from again

        except Exception as e:
            print(f'Unexpected error: {e}')
            set_status('An error occurred. Stopping.')
            time.sleep(3)
            running = False


# ── GUI helpers ───────────────────────────────────────────────────────────────

def set_status(text):
    result_label.config(text=text, font=('Comfortaa', 12, 'bold'))
    root.update()


def start_script():
    global running
    if running:
        return
    confirmed = pyautogui.confirm(text='Start the automator?', title='Confirmation', buttons=['OK', 'Cancel'])
    if confirmed == 'Cancel':
        return
    running = True
    threading.Thread(target=run_automation, daemon=True).start()


def stop_script():
    global running
    running = False
    set_status('Stopped.')
    print('Stopped.')


# ── GUI layout ────────────────────────────────────────────────────────────────

root = Tk()
root.geometry('250x280')
root.attributes('-topmost', True)

Label(root, text='Time Attendance Automator', font=('Comfortaa', 10, 'bold')).pack(pady=10)

Button(root, text='Start Automation', command=start_script).pack(pady=5)
Button(root, text='Stop Automation',  command=stop_script).pack(pady=5)
Button(root, text='Quit',             command=root.destroy).pack(pady=5)

result_label = Label(root, text='', wraplength=220)
result_label.pack(pady=10)

root.mainloop()