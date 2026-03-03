import os
import time
import subprocess
import pyautogui

# === CONFIG ===
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
PROFILE_DIR = r"C:\Users\therm\AppData\Local\Google\Chrome\User Data\Profile 1"
EXPORT_WAIT_TIME = 45  # seconds for "Preparing PDF..."

report_links = [
    "https://lookerstudio.google.com/u/0/reporting/f512fb52-497b-4afb-846a-7a8259cb9e0d/page/p_8hcedim9jd/edit",
    "https://lookerstudio.google.com/u/0/reporting/b23e56ae-2b75-4b49-a32f-f463a30cc7e5/page/p_ly32og54rd/edit"
]

def open_chrome_tabs():
    print("🚀 Opening Chrome tabs with Profile 1...")
    for url in report_links:
        subprocess.Popen([
            CHROME_PATH,
            f'--profile-directory={os.path.basename(PROFILE_DIR)}',
            '--start-maximized',
            '--new-window',
            url
        ])
        time.sleep(3)
    print("✅ All tabs opened.\n")

def export_current_tab_as_pdf():
    print("📄 Exporting current tab to PDF...")

    # Step 1: Go fullscreen
    pyautogui.press('f11')
    time.sleep(2)

    # Step 2: Click the "File" menu
    pyautogui.moveTo(70, 48)  # Your confirmed position
    pyautogui.click()
    time.sleep(1)

    # Step 3: Arrow to “Download as” > PDF
    pyautogui.press('down', presses=7, interval=0.2)
    pyautogui.press('right')
    pyautogui.press('enter')
    time.sleep(3)

    # Step 4: Press Enter on Download popup
    pyautogui.press('enter')
    print("⏳ Waiting for PDF to prepare...")
    time.sleep(EXPORT_WAIT_TIME)

    # Step 5: Confirm Save in download dialog
    pyautogui.press('enter')
    print("💾 PDF saved.\n")
    time.sleep(3)

    # Step 6: Exit fullscreen
    pyautogui.press('f11')
    time.sleep(1)

def run_export_flow():
    print("⏳ Waiting 30 sec for first tab to load...")
    time.sleep(30)

    for i in range(len(report_links)):
        export_current_tab_as_pdf()
        if i < len(report_links) - 1:
            pyautogui.hotkey('ctrl', 'tab')
            print("➡️ Switching to next tab...")
            time.sleep(5)

if __name__ == "__main__":
    open_chrome_tabs()
    run_export_flow()
    print("🎉 All exports done!")
