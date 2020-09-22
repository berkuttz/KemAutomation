import os
import msedge.selenium_tools
import shutil
import time

options = msedge.selenium_tools.EdgeOptions()
options.use_chromium = True
if not os.path.isdir("C:\\Users\\d4an\\AppData\\Local\\Microsoft\\Edge\\User Data2"):
    print("Please, make sure MS Edge is closed")
    time.sleep(5)
    shutil.copytree("C:\\Users\\d4an\\AppData\\Local\\Microsoft\\Edge\\User Data",
                "C:\\Users\\d4an\\AppData\\Local\\Microsoft\\Edge\\User Data2")

options.add_argument("user-data-dir=C:\\Users\\d4an\\AppData\\Local\\Microsoft\\Edge\\User Data2")
driver = msedge.selenium_tools.Edge(executable_path=r'C:\Users\d4an\Downloads\msedgedriver.exe', options=options)
driver.maximize_window()
