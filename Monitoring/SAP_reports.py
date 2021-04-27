import win32com.client
from datetime import datetime

print("this")

print("2")

# check if SAP window is opended
def check_sap():
    SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    SessionNr = 6
    listOfOpenSess = []

    # Check what sessions are open and get list of openned sessions
    for i in range(SessionNr)[::-1]:
        try:
            SapGui.FindById("ses[" + str(i) + "]")
            listOfOpenSess.append(i)
        except:
            continue
    listOfOpenSess.sort()
    if len(listOfOpenSess) == 0:
        return False
    else:
        return True


if check_sap():
    weekday = datetime.today().weekday()
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(r"C:\Users\d4an\OneDrive - Kemira Oyj\Desktop\Monitoring.xlsm", ReadOnly=1)
    xl.Application.Run("Monitoring.xlsm!VA14L")
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl

with open("report_log.txt", "w"):
    pass