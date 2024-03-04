from tkinter.filedialog import askopenfilename
import pandas as pd
import time
import win32gui
import win32com.client
import subprocess
from datetime import datetime
import os

from openpyxl.workbook import Workbook
from pynput.keyboard import Controller
import datetime as dt

## Setting variables to check is this version matches with the GSS Automation Team's control
application = "Month-End Close De Beers"
version = "v01"
user_name = os.getlogin()
path = f"C:/Users/{user_name}/Box/Automation Script Versions/versions.xlsx"
df = pd.read_excel(path)
filter_criteria = (df['app'] == application) & (df['versão'] == version)
start_time = None

if not filter_criteria.any():
    print('Outdated app, talk to the automation team. Press ENTER to close the code \n')
    quit()

userInput = input("Please press ENTER and select the file with valid company codes:\n")
file_name = askopenfilename()
folder_path = file_name[0:file_name.rfind("/") + 1]
current_month = input("Enter your current month:\n")
new_month = str(int(current_month) + 1).zfill(2)
today = dt.datetime.now().date()
next_year = today.year + 1
current_year = today.year
print(folder_path)

# Read closure data from an Excel template.
codes = pd.read_excel(file_name)

# SAP Scripting - to log in and navigate to the correct transaction.
keyboard = Controller()

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

# open the SAP
print("Opening SAP")
time.sleep(1)
path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
subprocess.Popen(path)
time.sleep(0.5)
# Connect to SAP GUI
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine

# Ask the user for the server name
server_name = input("Please enter your SAP server name: \n")

if server_name == "QP8":
    # Open connection
    connection = application.OpenConnection(str(server_name), True) # "QP8"

    # Create session
    session = connection.Children(0)
    session.findById("wnd[0]").maximize()
    username = input("Please enter your SAP username \n")
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
    password = input("Please enter your SAP password \n")
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password

    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "OB52"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[5]").press()

elif server_name == "PS8 [Anglo AOP]":
    # Open connection
    connection = application.OpenConnection(str(server_name), True)

    # Create session
    session = connection.Children(0)
    session.findById("wnd[0]").maximize()
    # username = input("Please enter your SAP username \n")
    # session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
    # password = input("Please enter your SAP password \n")
    # session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password

    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "OB52"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[5]").press()

else:
    print("Invalid server name. Please enter either QP8 or PS8")

def run_first_loop(codes, new_month):
    df_sap = pd.DataFrame(columns=["Var.",	"A",	"From Account",	"To Account",	"From Per. 1",	"Year",	"To Per. 1",	"Year2",	"Authorization Group",	"From Per.2",	"Year3",	"To Per. 2",	"Year4", "Status", "Date_Time"])
    wb = Workbook()
    now = dt.datetime.now().date()
    file_path = r'C:\Users\Public\Documents\Month-End Close De Beers\Reports\report'+str(now)+'.xlsx'
    wb.save(file_path)

    _sbar = session.FindById("wnd[0]/sbar/pane[0]").Text
    while _sbar == "" or _sbar[:4] == "Peri":
        if _sbar == "":
            actual_code = session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-BUKRS[0,0]").text
            if session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text == 'M':
                if (codes == actual_code).any().any():
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text = new_month
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text = new_month
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text = new_month
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text = new_month

                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text = current_year
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text = current_year
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text = current_year
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text = current_year

                    if current_month == "12":
                        # Roll forward to the next year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text = "01"
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text = "01"
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text = "01"
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text = "01"

                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text = next_year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text = next_year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text = next_year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text = next_year

                    df_sap.loc[len(df_sap)] = {
                        'Var.': actual_code,
                        'A': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text,
                        'From Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-VKONT[2,0]").text,
                        'To Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BKONT[3,0]").text,
                        'From Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text,
                        'Year': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text,
                        'To Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text,
                        'Year2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text,
                        'Authorization Group': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BRGRU[8,0]").text,
                        'From Per.2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text,
                        'Year3': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text,
                        'To Per. 2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text,
                        'Year4': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text,
                        'Status': "Closed",
                        'Date_Time': pd.Timestamp.now()}
                else:
                    df_sap.loc[len(df_sap)] = {
                        'Var.': actual_code,
                        'A': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text,
                        'From Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-VKONT[2,0]").text,
                        'To Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BKONT[3,0]").text,
                        'From Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text,
                        'Year': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text,
                        'To Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text,
                        'Year2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text,
                        'Authorization Group': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BRGRU[8,0]").text,
                        'From Per.2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text,
                        'Year3': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text,
                        'To Per. 2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text,
                        'Year4': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text,
                        'Status': "Code not on the company code list",
                        'Date_Time': pd.Timestamp.now()}
            else:
                df_sap.loc[len(df_sap)] = {
                    'Var.': actual_code,
                    'A': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text,
                    'From Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-VKONT[2,0]").text,
                    'To Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BKONT[3,0]").text,
                    'From Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text,
                    'Year': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text,
                    'To Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text,
                    'Year2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text,
                    'Authorization Group': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BRGRU[8,0]").text,
                    'From Per.2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text,
                    'Year3': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text,
                    'To Per. 2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text,
                    'Year4': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text,
                    'Status': "Skipped because it is not an M",
                    'Date_Time': pd.Timestamp.now()}

        elif _sbar[:4] == "Peri":
            df_sap.loc[len(df_sap)] = {'Status': "Blocked Company code"}
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        session.findById("wnd[0]/mbar/menu[2]/menu[2]").select()
        _sbar = session.FindById("wnd[0]/sbar/pane[0]").Text


    with pd.ExcelWriter(r'C:\Users\Public\Documents\Month-End Close De Beers\Reports\report'+str(now)+'.xlsx', mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
        df_sap.to_excel(writer, sheet_name="report", index=False, header=True, startrow=0)
        print("Completed successfully")



def run_second_loop(codes, new_month):
    df_sap = pd.DataFrame(
        columns=["Var.", "A", "From Account", "To Account", "From Per. 1", "Year", "To Per. 1", "Year2", "Authorization Group", "From Per.2", "Year3",    "To Per. 2", "Year4", "Status", "Date_Time"])
    wb = Workbook()
    now = dt.datetime.now().date()
    file_path = r'C:\Users\Public\Documents\Month-End Close De Beers\Reports\report'+str(now)+'.xlsx'
    wb.save(file_path)

    _sbar = session.FindById("wnd[0]/sbar/pane[0]").Text
    while _sbar == "" or _sbar[:4] == "Peri":
        if _sbar == "":
            actual_code = session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-BUKRS[0,0]").text
            if session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text != 'M':
                if (codes == actual_code).any().any():

                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text = new_month
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text = new_month
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text = new_month
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text = new_month

                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text = current_year
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text = current_year
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text = current_year
                    session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text = current_year

                    if current_month == "12":
                        # Roll forward to the next year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text = "01"
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text = "01"
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text = "01"
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text = "01"

                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text = next_year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text = next_year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text = next_year
                        session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text = next_year

                    df_sap.loc[len(df_sap)] = {
                        'Var.': actual_code,
                        'A': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text,
                        'From Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-VKONT[2,0]").text,
                        'To Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BKONT[3,0]").text,
                        'From Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text,
                        'Year': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text,
                        'To Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text,
                        'Year2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text,
                        'Authorization Group': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BRGRU[8,0]").text,
                        'From Per.2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text,
                        'Year3': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text,
                        'To Per. 2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text,
                        'Year4': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text,
                        'Status': "Closed",
                        'Date_Time': pd.Timestamp.now()}
                else:
                    df_sap.loc[len(df_sap)] = {
                        'Var.': actual_code,
                        'A': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text,
                        'From Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-VKONT[2,0]").text,
                        'To Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BKONT[3,0]").text,
                        'From Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text,
                        'Year': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text,
                        'To Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text,
                        'Year2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text,
                        'Authorization Group': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BRGRU[8,0]").text,
                        'From Per.2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text,
                        'Year3': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text,
                        'To Per. 2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text,
                        'Year4': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text,
                        'Status': "Code not on the company code list",
                        'Date_Time': pd.Timestamp.now()}

            else:
                df_sap.loc[len(df_sap)] = {
                    'Var.': actual_code,
                    'A': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/ctxtV_T001B_COFI-MKOAR[1,0]").text,
                    'From Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-VKONT[2,0]").text,
                    'To Account': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BKONT[3,0]").text,
                    'From Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE1[4,0]").text,
                    'Year': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE1[5,0]").text,
                    'To Per. 1': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE1[6,0]").text,
                    'Year2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE1[7,0]").text,
                    'Authorization Group': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-BRGRU[8,0]").text,
                    'From Per.2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRPE2[9,0]").text,
                    'Year3': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-FRYE2[10,0]").text,
                    'To Per. 2': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOPE2[11,0]").text,
                    'Year4': session.findById("wnd[0]/usr/tblSAPL0F00TCTRL_V_T001B_COFI/txtV_T001B_COFI-TOYE2[12,0]").text,
                    'Status': "Skipped because it is M",
                    'Date_Time': pd.Timestamp.now()}

        elif _sbar[:4] == "Peri":
            df_sap.loc[len(df_sap)] = {'Status': "Blocked Company code"}
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        session.findById("wnd[0]/mbar/menu[2]/menu[2]").select()
        _sbar = session.FindById("wnd[0]/sbar/pane[0]").Text

    with pd.ExcelWriter(r'C:\Users\Public\Documents\Month-End Close De Beers\Reports\report'+str(now)+'.xlsx', mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
        df_sap.to_excel(writer, sheet_name="report", index=False, header=True, startrow=0)
        print("Completed successfully")


# Read the schedule Excel file
schedule_data = pd.read_excel('Annual Calendar.xlsx')

# It converts the 'Date' column in the schedule_data Excel file to a datetime format.
schedule_data['Date'] = pd.to_datetime(schedule_data['Date'])

# Read the schedule Excel file
schedule_data = pd.read_excel('Annual Calendar.xlsx')

# It converts the 'Date' column in the schedule_data Excel file to a datetime format.
schedule_data['Date'] = pd.to_datetime(schedule_data['Date'])

# It iterates through each row in the Annual Calendar Excel file and for each row, it extracts the date and compares it to today's date.
for index, row in schedule_data.iterrows():
    schedule_date = row['Date']
    account_type = row['Account Type']
    # Check for specific dates and execute corresponding SAP automation logic
    if schedule_date.date() == today:
        # Check if account_type is M
        if account_type == 'M':
            print("Closing all 'M' items.....")
            run_first_loop(codes, new_month)
        elif schedule_date.date() == today:
            if account_type == 'Other':
                print("Closing all items except 'M' items.....")
                run_second_loop(codes, new_month)
        else:
            print("We're skipping today's execution because it's not included on our yearly calendar. If you would like to execute today, please update our annual calendar.")

# Function to create the file path and directory if it doesn't exist
def create_file_path():
    now = datetime.now().date()
    folder_path = r'C:\Users\Public\Documents\Month-End Close De Beers\Reports'
    file_path = fr'{folder_path}\report{str(now)}.xlsx'

    try:
        # Check if the directory exists or create it
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        return file_path
    except OSError as e:
        print(f"Error creating directory: {e}")
        return None

# Get the file path
file_path = create_file_path()

if file_path:
    print(f"File path: {file_path}")
else:
    print("File path creation failed.")

################################ LOG PREPARATION ##################################

# Function to create the file path and directory if it doesn't exist
def create_file_path(file_name):
    now = datetime.now().date()
    folder_path = r'C:\Users\Public\Documents\Month-End Close De Beers'
    file_path = os.path.join(folder_path, file_name)

    try:
        # Check if the directory exists or create it
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Create the path for the LogControl folder
        log_folder_path = os.path.join(folder_path, "LogControl")
        if not os.path.exists(log_folder_path):
            os.makedirs(log_folder_path)

        # Create the full path to the log file within the LogControl folder
        log_file_path = os.path.join(log_folder_path, f"execution_log_{now}.txt")
        return file_path, log_file_path
    except OSError as e:
        print(f"Error creating directory: {e}")
        return None, None

# Get the file path and log file path
file_name = 'report' + str(datetime.now()) + '.xlsx'
file_path, log_file_path = create_file_path(file_name)

if file_path and log_file_path:
    print(f"File path: {file_path}")
    print(f"Log file path: {log_file_path}")
else:
    print("File path creation failed.")

# Save the execution log
def save_execution_log(log_file_path, message):
    with open(log_file_path, 'a') as log_file:
        log_file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")

# Once the period has been updated, save.
save_execution_log(log_file_path, "Period updated successfully.")
# Once the period has been updated, save.
# session.findById("wnd[0]/tbar[0]/btn[11]").press()


