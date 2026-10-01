"""
Daily Status Report Bot
-----------------------
Visible, step-by-step desktop automation:

  Win+R -> type "chrome" -> Enter -> Ctrl+L -> type URL -> Enter
  -> Ctrl+A -> Ctrl+C (weather text)
  Win+R -> type "excel"  -> Enter -> Ctrl+N (blank workbook)
  -> paste headers + report row -> format -> save -> screenshot
"""

import os
import time
import urllib.request
from datetime import datetime

import pyautogui
import pyperclip
import pywintypes
import win32com.client
import win32con
import win32gui


# ---------------------------------------------------------
# Configuration
# ---------------------------------------------------------

WEATHER_URL = "https://wttr.in/Chennai?format=3"
CITY = "Chennai"

SAVE_FOLDER = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "Daily_Report_Bot"
)

os.makedirs(SAVE_FOLDER, exist_ok=True)

pyautogui.FAILSAFE = True      # move mouse to top-left corner to abort
pyautogui.PAUSE = 0.5


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

def step(number, text):
    print(f"\n[STEP {number}] {text}")


# Window class names - more reliable than titles, which can match
# the wrong window (e.g. a terminal or browser tab mentioning "Excel").
RUN_DIALOG_CLASS = "#32770"
CHROME_CLASS = "Chrome_WidgetWin_1"
EXCEL_CLASS = "XLMAIN"


def find_window(title_part="", class_name=None, timeout=20, exact=False):
    """Wait until a visible window matching the class/title appears."""
    end = time.time() + timeout
    while time.time() < end:
        found = []

        def callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            if class_name and win32gui.GetClassName(hwnd) != class_name:
                return
            title = win32gui.GetWindowText(hwnd)
            if exact:
                if title == title_part:
                    found.append(hwnd)
            elif title_part.lower() in title.lower():
                found.append(hwnd)

        win32gui.EnumWindows(callback, None)
        if found:
            return found[0]
        time.sleep(0.5)
    return None


def bring_to_front(hwnd):
    """Maximize a window and force it to the foreground."""
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        # Windows only allows SetForegroundWindow after recent input.
        # Use Shift, NOT Alt: a lone Alt press activates menu mode and
        # swallows the next key (that's what turned "chrome" into "hrome").
        pyautogui.press("shift")
        win32gui.SetForegroundWindow(hwnd)
    except Exception as e:
        print(f"    Warning: could not focus window ({e})")
    time.sleep(1)


def open_with_run_dialog(program):
    """Start -> Run -> type program -> Enter."""
    print("    Pressing Win+R to open the Run dialog...")
    pyautogui.hotkey("win", "r")

    run_hwnd = find_window("Run", class_name=RUN_DIALOG_CLASS,
                           timeout=5, exact=True)
    if run_hwnd:
        bring_to_front_no_max(run_hwnd)
    else:
        time.sleep(1)

    time.sleep(0.5)
    # Run remembers the last command - select it so typing replaces it
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")

    print(f"    Typing '{program}'...")
    pyautogui.write(program, interval=0.05)

    print("    Pressing Enter...")
    pyautogui.press("enter")


def bring_to_front_no_max(hwnd):
    """Focus a small dialog without maximizing it."""
    try:
        if win32gui.GetForegroundWindow() != hwnd:
            pyautogui.press("shift")
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.5)


def connect_to_excel(excel_hwnd, timeout=30):
    """
    Attach to the Excel instance that was opened via Win+R.

    Excel only registers itself for GetActiveObject after its window
    loses focus once, so we nudge focus away and back while retrying.
    """
    end = time.time() + timeout
    while time.time() < end:
        try:
            return win32com.client.GetActiveObject("Excel.Application")
        except pywintypes.com_error:
            pyautogui.hotkey("alt", "tab")      # move focus away...
            time.sleep(1)
            bring_to_front(excel_hwnd)          # ...and back
    return None


# ---------------------------------------------------------
# Date / Time and file names
# ---------------------------------------------------------

now = datetime.now()

current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")
current_date = now.strftime("%Y-%m-%d")

excel_path = os.path.join(SAVE_FOLDER, f"daily_report_{current_date}.xlsx")
screenshot_path = os.path.join(SAVE_FOLDER, f"daily_report_{current_date}.png")


print("=" * 60)
print("Daily Status Report Automation Started")
print("=" * 60)

excel = None

try:

    # -----------------------------------------------------
    # STEP 1 - Start -> Run -> chrome
    # -----------------------------------------------------

    step(1, "Opening Google Chrome (Start -> Run -> chrome)")

    open_with_run_dialog("chrome")

    chrome_hwnd = find_window("Google Chrome", class_name=CHROME_CLASS,
                              timeout=20)
    if not chrome_hwnd:
        raise Exception("Chrome window did not appear.")

    bring_to_front(chrome_hwnd)
    print("    Chrome is open and focused.")


    # -----------------------------------------------------
    # STEP 2 - Type the weather URL in the address bar
    # -----------------------------------------------------

    step(2, "Opening weather website")

    print("    Pressing Ctrl+L (address bar)...")
    pyautogui.hotkey("ctrl", "l")

    print(f"    Typing {WEATHER_URL}...")
    pyautogui.write(WEATHER_URL, interval=0.03)

    print("    Pressing Enter...")
    pyautogui.press("enter")

    time.sleep(6)


    # -----------------------------------------------------
    # STEP 3 - Copy the weather text (Ctrl+A, Ctrl+C)
    # -----------------------------------------------------

    step(3, "Copying weather information")

    # Make sure Chrome (not the code editor) has focus before copying
    bring_to_front(chrome_hwnd)

    pyperclip.copy("")            # clear clipboard so old text isn't reused

    # Click inside the page so Ctrl+A selects page text, not the address bar
    screen_w, screen_h = pyautogui.size()
    pyautogui.click(screen_w // 2, screen_h // 2)

    print("    Pressing Ctrl+A...")
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)

    print("    Pressing Ctrl+C...")
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)

    weather_data = " ".join(pyperclip.paste().strip().split())


    # -----------------------------------------------------
    # STEP 4 - Validate (fallback to direct download)
    # -----------------------------------------------------

    step(4, "Validating weather data")

    if CITY not in weather_data or len(weather_data) > 100:
        print("    Copied text doesn't look like weather data - "
              "fetching directly instead.")
        request = urllib.request.Request(
            WEATHER_URL, headers={"User-Agent": "curl/8.0"}
        )
        with urllib.request.urlopen(request, timeout=20) as response:
            weather_data = " ".join(
                response.read().decode("utf-8").strip().split()
            )

    if CITY not in weather_data:
        raise Exception(f"Could not get weather data: {weather_data[:100]!r}")

    comment = "Weather information collected successfully."

    print(f"    Fetched Data: {weather_data}")
    print(f"    Comment     : {comment}")


    # -----------------------------------------------------
    # STEP 5 - Start -> Run -> excel
    # -----------------------------------------------------

    step(5, "Opening Microsoft Excel (Start -> Run -> excel)")

    open_with_run_dialog("excel")

    excel_hwnd = find_window(class_name=EXCEL_CLASS, timeout=30)
    if not excel_hwnd:
        raise Exception("Excel window did not appear.")

    time.sleep(3)                  # let the Excel start screen finish loading
    bring_to_front(excel_hwnd)
    print("    Excel is open and focused.")


    # -----------------------------------------------------
    # STEP 6 - Create a blank workbook (Ctrl+N)
    # -----------------------------------------------------

    step(6, "Creating new blank workbook (Ctrl+N)")

    pyautogui.hotkey("ctrl", "n")
    time.sleep(4)

    excel_hwnd = find_window(class_name=EXCEL_CLASS, timeout=10) or excel_hwnd
    bring_to_front(excel_hwnd)


    # -----------------------------------------------------
    # STEP 7 - Connect to the open Excel (for save/format)
    # -----------------------------------------------------

    step(7, "Connecting to the open Excel window")

    excel = connect_to_excel(excel_hwnd)

    if excel is None:
        # Happens when Python runs as Administrator and Excel doesn't
        # (or vice versa). Fall back to an Excel we start ourselves.
        print("    Could not attach to that Excel window "
              "(check that Python and Excel run at the same privilege level).")
        print("    Starting a new Excel instance instead...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.Workbooks.Add()
        excel_hwnd = excel.Hwnd

    if excel.Workbooks.Count == 0:
        excel.Workbooks.Add()

    excel.DisplayAlerts = False
    workbook = excel.ActiveWorkbook
    worksheet = workbook.ActiveSheet

    bring_to_front(excel_hwnd)
    worksheet.Range("A1").Select()
    print("    Connected.")


    # -----------------------------------------------------
    # STEP 8 - Paste headers into row 1
    # -----------------------------------------------------

    step(8, "Adding report headers (A1:C1)")

    pyperclip.copy("Date & Time\tWeather Data\tComment")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)


    # -----------------------------------------------------
    # STEP 9 - Paste report row into row 2
    # -----------------------------------------------------

    step(9, "Adding daily report row (A2:C2)")

    worksheet.Range("A2").Select()
    time.sleep(0.5)

    pyperclip.copy(f"{current_datetime}\t{weather_data}\t{comment}")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    pyautogui.press("esc")        # leave paste mode
    time.sleep(1)


    # -----------------------------------------------------
    # STEP 10 - Formatting
    # -----------------------------------------------------

    step(10, "Formatting sheet (bold headers, auto-fit columns)")

    worksheet.Range("A1:C1").Font.Bold = True
    worksheet.Columns("A:C").AutoFit()


    # -----------------------------------------------------
    # STEP 11 - Save workbook
    # -----------------------------------------------------

    step(11, "Saving Excel report")

    # If today's report is still open from an earlier run it's locked;
    # save under a time-stamped name instead of crashing.
    if os.path.exists(excel_path):
        try:
            os.remove(excel_path)
        except PermissionError:
            stamp = now.strftime("%H%M%S")
            excel_path = os.path.join(
                SAVE_FOLDER, f"daily_report_{current_date}_{stamp}.xlsx")
            screenshot_path = os.path.join(
                SAVE_FOLDER, f"daily_report_{current_date}_{stamp}.png")
            print("    Today's report is open elsewhere (locked) - "
                  f"saving as {os.path.basename(excel_path)}")

    workbook.SaveAs(excel_path, FileFormat=51)     # 51 = .xlsx
    time.sleep(2)

    if not os.path.exists(excel_path):
        raise Exception(f"Excel file was not created: {excel_path}")

    print(f"    Saved: {excel_path}")


    # -----------------------------------------------------
    # STEP 12 - Screenshot of the final sheet
    # -----------------------------------------------------

    step(12, "Taking screenshot of final Excel sheet")

    bring_to_front(excel_hwnd)
    worksheet.Range("A1").Select()
    time.sleep(2)

    pyautogui.screenshot().save(screenshot_path)

    if not os.path.exists(screenshot_path):
        raise Exception("Screenshot could not be saved.")

    print(f"    Saved: {screenshot_path}")


    # -----------------------------------------------------
    # Completed
    # -----------------------------------------------------

    print()
    print("=" * 60)
    print("Daily Report Created Successfully")
    print("=" * 60)
    print(f"Date & Time : {current_datetime}")
    print(f"Weather     : {weather_data}")
    print(f"Excel File  : {excel_path}")
    print(f"Screenshot  : {screenshot_path}")


except Exception as error:

    print()
    print("=" * 60)
    print("Automation failed.")
    print("=" * 60)
    print(f"Error: {error}")

finally:
    if excel is not None:
        try:
            excel.DisplayAlerts = True
        except Exception:
            pass