#Code writer gadhavek97@gmail.com

import  selenium
import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from tkinter import filedialog
import tkinter.messagebox
import tkinter as tk
import pyodbc
import zipfile
import datetime
import numpy as np
import os
import win32com.client
access = win32com.client.Dispatch('Access.Application')
from comtypes.gen import Access

root = tk.Tk()
root.title("Please select below option to proceed further!")
var = 0
def submit():
    global var, vars
    var = vars.get()
    if vars.get() == 1:
        print("Option: {var} selected.")
    elif vars.get() == 2:
        print(f"Option: {var} selected.")
    elif vars.get() == 3:
        print(f"Option: {var} selected.")
    else:
        print("Please select right option")
        exit()
    root.destroy()
vars = tk.IntVar()
checkbox1 = tk.Radiobutton(root,text="I want to download eCCI file and proceed further for compare report.", variable=vars, value=1, font=("Helvetica", 14))
checkbox2 = tk.Radiobutton(root,text="I have eCCI zip file and want to proceed further for compare report.", variable=vars, value=2, font=("Helvetica", 14))
checkbox3 = tk.Radiobutton(root,text="I want to download eCCI zip file only.", variable=vars, value=3, font=("Helvetica", 14))
submit_button = tk.Button(root, text="Submit", command=submit, font=("Helvetica", 14))
checkbox1.pack()
checkbox2.pack()
checkbox3.pack()
submit_button.pack()

for i in range(4):
    root.grid_rowconfigure(i, weight=2)
for i in range(2):
    root.grid_columnconfigure(i, weight=2)
window_width = 700
window_height = 160
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
root.mainloop()

if var in [1,3]:
    datalakeUser_id = None
    datalakePassword = None
    mainframeUser_id = None
    mainframePassword = None
    srNumber = None
    regionCode = None
    file_path = None

    def  open_csv_file():
        global file_path
        file_pat = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        file_path.set(file_pat)
    def submit_form():
        global datalakeUser_id, datalakePassword, mainframePassword, mainframeUser_id, srNumber, regionCode,file_path
        datalake_user_id = enter_datalake_user_id.get()
        datalake_password = enter_datalake_password.get()
        mainframe_user_id = enter_mainframe_user_id.get()
        mainframe_password = enter_mainframe_password.get()
        sr_number = enter_sr_number.get()
        region_code = enter_region_code.get()
        file_path = file_path_entery.get()

        datalakeUser_id = datalake_user_id
        datalakePassword = datalake_password
        mainframeUser_id = mainframe_user_id
        mainframePassword = mainframe_password
        srNumber = sr_number
        regionCode = region_code
        file_path = file_path
        root.destroy()
    root = tk.Tk()
    root.title('Please Enter The Required Details To Proceed')

    label_datalake_user_id = tk.Label(root, text='Data Lake User ID: ', font=("Helvetica", 14))
    label_datalake_user_id.grid(row=0, column=0)
    label_datalake_password = tk.Label(root, text='Data Lake Password: ', font=("Helvetica", 14))
    label_datalake_password.grid(row=1, column=0)
    label_mainframe_user_id = tk.Label(root, text='Mainframe User ID: ', font=("Helvetica", 14))
    label_mainframe_user_id.grid(row=2, column=0)
    label_mainframe_password = tk.Label(root, text='Mainframe Password: ', font=("Helvetica", 14))
    label_mainframe_password.grid(row=3, column=0)
    label_sr_number = tk.Label(root, text='Client Name: ', font=("Helvetica", 14))
    label_sr_number.grid(row=4, column=0)
    label_region_code = tk.Label(root, text='Region Name: ', font=("Helvetica", 14))
    label_region_code.grid(row=5, column=0)

    enter_datalake_user_id = tk.Entry(root, font=("Helvetica", 14))
    enter_datalake_user_id.grid(row=0, column=1)
    enter_datalake_password = tk.Entry(root, show="*", font=("Helvetica", 14))
    enter_datalake_password.grid(row=1, column=1)
    enter_mainframe_user_id = tk.Entry(root, font=("Helvetica", 14))
    enter_mainframe_user_id.grid(row=2, column=1)
    enter_mainframe_password = tk.Entry(root, show="*", font=("Helvetica", 14))
    enter_mainframe_password.grid(row=3, column=1)
    enter_sr_number = tk.Entry(root, font=("Helvetica", 14))
    enter_sr_number.grid(row=4, column=1)
    enter_region_code = tk.Entry(root, font=("Helvetica", 14))
    enter_region_code.grid(row=5, column=1)

    file_path = tk.StringVar()
    file_path_lable = tk.Label(root,text='Select Company Code CSV File', font=("Helvetica", 14))
    file_path_lable.grid(row=6, column=0)
    file_path_entery = tk.Entry(root, textvariable=file_path, state="readonly", font=("Helvetica", 14))
    file_path_entery.grid(row=6, column=1)
    file_button = tk.Button(root, text='Browse File', command= open_csv_file, font=("Helvetica", 14))
    file_button.grid(row=6, column=2, pady=10,padx=20)
    submit_button = tk.Button(root, text="Submit", command=submit_form, font=("Helvetica", 14))
    submit_button.grid(row=7, columnspan=4, pady=10, padx=20)

    for i in range(7):
        root.grid_rowconfigure(i, weight=1)
    for i in range(2):
        root.grid_columnconfigure(i, weight=1)
    window_width = 700
    window_height = 550
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    root.mainloop()

    #file_path = tk.filedialog.askopenfilename(title="Please select csv formatted Company Code file..")
    csv_file_name = os.path.basename(file_path)
    lengths = len(csv_file_name)
    lths = len(file_path)
    pathOfFolderOutput = file_path[:lths-lengths]
    if regionCode == '':
        print('Region Name should not be Empty!! Try again with Region Name')
        exit()
    try:
        try:
            driver = webdriver.Edge()
            print('Microsoft Edge Driver Connected..')
            driver.maximize_window()
        except:
            driver = webdriver.Chrome()
            print('Chrome Driver Connected..')
            driver.maximize_window()
    except:
        print('Please check your Microsoft Edge Driver version and try again.')

    df = pd.read_csv(file_path, dtype='object')
    coCodeList = list(df.iloc[0].values.astype(str))
    coCodeStr = ''.join(coCodeList)
    print(coCodeStr)

# govind kushwaha
    try:

        driver.get("https://etoolsprod.adp.com/")

    except:
        print('https://etoolsprod.adp.com/ is not reachable please check it and try again..')

    driver.implicitly_wait(90)

    WebDriverWait(driver, 90).until(
        EC.visibility_of_element_located((By.ID, 'login-form_username'))
    )
    driver.find_element(By.ID, 'login-form_username').send_keys(datalakeUser_id)

    driver.find_element(By.ID, 'verifUseridBtn').click()

    driver.find_element(By.ID, 'login-form_password').send_keys(datalakePassword)

    driver.find_element(By.ID, 'signBtn').click()
    '''
    try:
        #WebDriverWait(driver, 5).until( EC.visibility_of_element_located((By.XPATH, "//div[@id= 'common_alert']")))
        time.sleep(5)
        while driver.find_element(By.XPATH, "//div[@id= 'common_alert']").text == "Your entry is not valid. Try again":
            def submit_form():
                global datalakeUser_id, datalakePassword
                datalake_user_id = enter_datalake_user_id.get()
                datalake_password = enter_datalake_password.get()
                datalakeUser_id = datalake_user_id
                datalakePassword = datalake_password
                root.destroy()
            root = tk.Tk()
            
            root.title('Wrong Password or UserID, Try Again')
            label_datalake_user_id = tk.Label(root, text='Data Lake User ID: ', font=("Helvetica", 14))
            label_datalake_user_id.grid(row=0, column=0)
            label_datalake_password = tk.Label(root, text='Data Lake Password: ', font=("Helvetica", 14))
            label_datalake_password.grid(row=1, column=0)

            enter_datalake_user_id = tk.Entry(root, font=("Helvetica", 14))
            enter_datalake_user_id.grid(row=0, column=1)
            enter_datalake_password = tk.Entry(root, show="*", font=("Helvetica", 14))
            enter_datalake_password.grid(row=1, column=1)
            submit_button = tk.Button(root, text="Submit", command=submit_form, font=("Helvetica", 14))
            submit_button.grid(row=7, columnspan=4, pady=10, padx=20)

            for i in range(2):
                root.grid_rowconfigure(i, weight=2)
            for i in range(2):
                root.grid_columnconfigure(i, weight=2)
            window_width = 600
            window_height = 200
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            x_position = (screen_width - window_width) // 2
            y_position = (screen_height - window_height) // 2
            root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
            root.mainloop()
            driver.find_element(By.ID, 'login-form_username').click()
            driver.find_element(By.ID, 'login-form_username').send_keys(Keys.CONTROL + "A")
            driver.find_element(By.ID, 'login-form_username').send_keys(Keys.DELETE)
            driver.find_element(By.ID, 'login-form_username').send_keys(datalakeUser_id)
            driver.find_element(By.ID, 'verifUseridBtn').click()

            driver.find_element(By.ID, 'login-form_password').clear()
            driver.find_element(By.ID, 'login-form_password').send_keys(datalakePassword)

            driver.find_element(By.ID, 'signBtn').click()
    except:
        print("Logged in Data Lake")
    '''
    driver.find_element(By.ID, 'widget_revit_form_TextBox_1').click()

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@id='revit_form_TextBox_1']"))
    )
    time.sleep(2)
    driver.find_element(By.XPATH, "//input[@id='revit_form_TextBox_1']").send_keys('Migrationtest')
    driver.find_element(By.XPATH, "//input[@id='revit_form_TextBox_1']").send_keys(Keys.ENTER)
    time.sleep(2)

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//td[@id='projectListingGrid_row_0_cell_0']//img[contains(@id,'projectListingGrid.store.rows')]"))
    )


    driver.find_element(By.XPATH, "//td[@id='projectListingGrid_row_0_cell_0']//img[contains(@id,'projectListingGrid.store.rows')]").click()

    driver.find_element(By.XPATH, "//span[@id='projectListingModuleEcciButton_label']").click()


    driver.get("https://etoolsprod.adp.com/eTools/landing.do")

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@id='revit_form_TextBox_1']"))
    )
    time.sleep(2)
    driver.find_element(By.XPATH, "//input[@id='revit_form_TextBox_1']").send_keys('Migrationtest')
    driver.find_element(By.XPATH, "//input[@id='revit_form_TextBox_1']").send_keys(Keys.ENTER)

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//td[@id='projectListingGrid_row_0_cell_0']//img[contains(@id,'projectListingGrid.store.rows')]"))
    )
    time.sleep(2)
    driver.find_element(By.XPATH, "//td[@id='projectListingGrid_row_0_cell_0']//img[contains(@id,'projectListingGrid.store.rows')]").click()

    driver.find_element(By.XPATH, "//span[@id='projectListingModuleEcciButton_label']").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//div[@class='rightButtons']//img[1]").click()
    driver.find_element(By.XPATH,"//input[@id='workbookName']").send_keys(srNumber)

    try:
        driver.find_element(By.XPATH,"//option[normalize-space()= '"+ regionCode +"' ]").click()
        values = driver.find_element(By.XPATH,"//option[normalize-space()= '"+ regionCode +"' ]")
    except:
        print(f'{regionCode} region Code did not find..')

    iframeVal = values.get_attribute('value')
    iframeValue = 'iframe' + str(iframeVal)

    driver.find_element(By.XPATH,"//img[@id='addRegion']").click()
    driver.execute_script("window.scrollTo(0, 500)")
    time.sleep(2)
    #li_ele = driver.find_elements(By.XPATH, f"//iframe[@name='{iframeValue}']")

    driver.switch_to.frame(driver.find_element(By.XPATH, f"//iframe[@name='{iframeValue}']"))

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//form/input[@id='fileToUpload']"))
    )

    driver.find_element(By.XPATH, "//form/input[@id='fileToUpload']").send_keys(file_path)

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@name='submit']"))
    )

    driver.find_element(By.XPATH, "//input[@name='submit']").click()

    driver.switch_to.default_content()
    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//option[@value='" + str(coCodeStr) + "']"))
    )
    try:
        pathOfCompanyCode = "//option[@value='" + str(coCodeStr) + "']"
        driver.find_element(By.XPATH,pathOfCompanyCode).click()
    except:
        print(f'{coCodeStr} Company Code did not find in Data Lake')

    action = ActionChains(driver)

    # perform the operation
    action.key_down(Keys.CONTROL).send_keys('A').key_up(Keys.CONTROL).perform()

    addbtn = "//img[@id='addCompanyCode"  + str(iframeVal) + "']"
    driver.find_element(By.XPATH, addbtn).click()


    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//img[@id='save_workbook']"))
    )
    driver.find_element(By.XPATH, "//img[@id='save_workbook']").click()

    WebDriverWait(driver, 60).until(
        EC.alert_is_present()
    )
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except:
        print('Alert did not found.. trying other way in 10 sec')
        action.send_keys(Keys.ENTER)

    driver.find_element(By.XPATH, "//a[@title='Data Queries']").click()

    driver.find_element(By.XPATH,"//input[@id='selectAll2']").click()

    driver.find_element(By.XPATH,"//a[@title='Data Queries for Data Compare']").click()

    driver.find_element(By.XPATH, "//input[@id='selectAll3']").click()

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//img[@id='generateQueryForDataCompareButton']"))
    )
    driver.find_element(By.XPATH, "//img[@id='generateQueryForDataCompareButton']").click()

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@id='cesnUserName']"))
    )
    try:
        driver.find_element(By.XPATH, "//input[@id='cesnUserName']").send_keys(mainframeUser_id)
        driver.find_element(By.XPATH, "//input[@id='cesnPassword']").send_keys(mainframePassword)
        driver.find_element(By.XPATH, "//input[@id='loginButton']").click()
    except:
        print('Mainframe password or ID is not correct')

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//div[@id='progressBox']"))
    )
    driver.find_element(By.XPATH, "//div[@id='progressBox']")

    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "//div[@id='progressText']"))
    )
    divText = driver.find_element(By.XPATH, "//div[@id='progressText']")
    checks = True

    t = 5
    val = 0
    while checks:
        try:
            divText = driver.find_element(By.XPATH, "//div[@id='progressText']")
            vals = divText.text[-10:-7]
            if vals != '':
                val = int(vals)
        except:
            pass
        if divText.text == 'Processing your request... 100% done.' or divText.text == '':
            checks = False
        if divText.text == 'Processing your request... 95% done.' or val > 90:
            t = 0
        time.sleep(t)
    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.XPATH, "(//td/font/a)[1]"))
    )
    time.sleep(1)
    driver.find_element(By.XPATH, "(//td/font/a)[1]").click()

    pathOfMainUsers = os.path.expanduser('~')
    pathOfDownloads = os.path.join(pathOfMainUsers, 'Downloads')
    #time.sleep(5)
    findZipFile = srNumber + '_DataQuery_' + str(regionCode).strip() + '_'
    chks = 0
    whilecheck = True
    def downloads_done():
        global chks, whilecheck
        chks = 1
        while whilecheck:
            for filenamess in os.listdir(pathOfDownloads):
                if filenamess.endswith(".crdownload"):
                    time.sleep(1)
                    downloads_done()
                if filenamess.startswith(findZipFile):
                    whilecheck = False
                    print("File downloaded")
    if chks == 0:
        downloads_done()
    time.sleep(2)
    driver.close()
    print(f'Download completed!, now wait for {srNumber}_Data_Compare Report.')
# govind kushwaha
if var == 1:
    pathOfMainUser = os.path.expanduser('~')
    pathOfDownload = os.path.join(pathOfMainUser, 'Downloads')
    findZipFile = srNumber + '_DataQuery_' + str(regionCode).strip() + '_'
    current_dates = str(datetime.datetime.now().strftime('%Y%m%d'))

    for file in os.listdir(pathOfDownload):
        dateOfZipFile = file[-16:-8]
        if file.startswith(findZipFile):
            ZipFilePath = os.path.join(pathOfDownload, file)
    if ZipFilePath == '':
        print('Zip File did not find..')
    file_path = ZipFilePath
    zip_file_name = os.path.basename(file_path)
    lengths = len(zip_file_name)
    lths = len(file_path)
    pathOfFolder = file_path[:lths-lengths]

    mdb_df = pd.DataFrame()

    delete_Cols = {
    'SCSUM' : ['RESDESC'],
    'SCFM' : ['PRIORITY'],
    'SCADJ' : ['LN', 'STEP'],
    'SCQU' : ['PRIORITY'],
    'LOCID' : ['PA32_RLO'],
    'MRP': ['TERM_DT'],
    'GRS' : ['AUTO_PAY', 'BY_PAY_GRP', 'FLD3', 'DED_FLD3', 'FLD4', 'DED_FLD4', 'FLD5', 'DED_FLD5', 'REDUCE_HR4', 'REDUCE_HR5', 'REDUCE_HR6', 'REDUCE_HR7', 'REDUCE_HR8', 'REDUCE_HR9', 'REDUCE_HR10', 'REDUCE_HR11', 'REDUCE_HR12', 'REDUCE_HR13', 'REDUCE_HR14', 'REDUCE_HR15', 'REDUCE_HR16', 'REDUCE_HR17', 'REDUCE_HR18', 'REDUCE_HR19', 'REDUCE_HR20', 'REDUCE_HR21']
    }

    file_rename_Mappings1 = {'TRCO': ['TST','CMP','QAL-PN','AC-STS','AC-CD','ACCUM DESCRIPTION','CLR','TAXOPT','FLD-SYM'],
                            'AAC': ['TST','CMP','CD', 'DESCRIPTION','CLR','TAXOPT','LDGR','INCEXC','SYMB','AMC','EXP','DWN'],
    'SCSUM' : ['TST','CMP','RES', 'MTH-NUM', 'MTH-DSC', 'CALC-STS', 'ALT-RES', 'CLC-LVL'],
    'SCFM' : ['TST','CMP','RES', 'MTH', 'LIN#', 'OPR-A', 'FLD-SYM-A', 'OPR-B', 'FLD-SYM-B', 'OPR-C', 'FLD-SYM-C', 'INT-RES'],
    'SCADJ' : ['TST','CMP','RSLT-SYMB-C', 'MTH-NUM', 'FLD-OR-RES', 'TYP', 'AMT', 'ACC-FLD-SYMB-C'],
    'SCQU' : ['TST','CMP','RES', 'MTH', 'QL1', 'FR1', 'TO1', 'CD1', 'VAL1', 'QL2', 'FR2', 'TO2', 'CD2', 'VAL2', 'QL3', 'FR3', 'TO3', 'CD3', 'VAL3'],
    'GTL' : ['TST','CMP','GTLTYPE', 'COD', 'TAXOPT', 'CLCOCR', 'STDHRS', 'INSFCT', 'ANNMAX', 'RNDNG', 'BDAY', 'EEDED'],
    'SCBA' : ['TST','CMP','RES', 'MTH', 'TRRST', 'CURACR', 'CD', 'LIT', 'TYP', 'BNTYP', 'RDTRDOL', 'TRN-TIM', 'TRICR', 'TRAMT', 'EXCTKN', 'UNALLAMT'],
    'SMHD' : ['TST','CMP','DPT-NUM', 'DPT-DSC'],
    'LOCID': ['TST','CMP','STATECD', 'LOCALTXCD', 'Expr1', 'LOCALNAM'],
    'GEN' : ['TST','CMP','VER', 'DEPTSIZ', 'QTR', 'CSR', 'ST-RCP', 'CTY-RCP', 'MJUR', 'PCS'],
    'GRS' : ['TST','CMP','REDSTHR1', 'REDSTHR2', 'REDSTHR3', 'SALOTHRS', 'CLCRT2', 'GRSCLC2', 'GRSCLC3', 'GRSCLC4', 'ALLUSE', 'STDHRS'],
    'TLM' : ['TST','CMP','TOTALTIME', 'EPIP',	'OTH_TLM'],
    'STUBL' : ['TST','CMP','FLD-SYM', 'LNG-DSC-COI-L1', 'LNG-DSC-COI-L2'],
    'PAYEC' : ['TST','CMP','OMNI#', 'PRINTAUDMSG', 'CHKSIGNING', 'STUFFSEAL'],
    'CUST' : ['TST','CMP','CUSTNETPROC', 'CUSTPROG', 'CUSTREPORTS', 'NOSCUSTIND', 'CUSTCHECK'],
    'CPCDA' : ['TST','CMP','CUSTOM', 'REPORT', 'CTRL', 'REC', 'SEQ#', 'CTRLCARDS'],
    'MRP' : ['TST','CMP','AUDRPT', 'WRKSHTSRT', 'MAJSRT', 'INTSRT', 'SRTBY', 'ACTFLGSEQ', 'TAXSVCRPT', 'NHIRECOMPL', 'HIREDT', 'BRTHDT', 'LOADT'],
    'RSCH' : ['TST','CMP','RSG', 'RUN-FRQ', 'DET-SAV', 'DELIVERY', 'FOR-PER-COD', 'WK-DAY-COD', 'STA-COD'],
    'PCAL' : ['TST','CMP','PAYFRQ', 'WK', 'PEDAT', 'INDAT', 'OUTDAT', 'PAYDAT'],
    'INET' : ['TST','CMP','IPAY', 'IREPORTS', 'SUP PAPER', 'PR QUICKVIEW', 'REMCTRL', 'SELF-SVC', 'W2_REISSUE', 'P_FREEVERS', 'P_CONTMAN', 'ORG_OID', 'RES_ADPREG', 'REALTIME']
                            }

    current_date = str(datetime.datetime.now().strftime('%d-%m-%Y'))


    if not os.path.exists(f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}'):
        os.makedirs(f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}')
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}')

    if os.path.exists(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}'):
        try:
            os.rmdir(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}')
        except OSError as e:
            print(f'Folder {srNumber}_Data_OutPut does not exist, creating new folder....')

    if not os.path.exists(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}'):
        os.makedirs(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}')
        os.makedirs(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}/{srNumber}_Text Files')

    New_file_path = f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}'

    try:

        DB_Engine = access.DBEngine
        DB_Engine.CreateDatabase(f'{pathOfFolderOutput}{srNumber}_Data_OutPut_{current_date}/DataCompare_Database.mdb', Access.DB_LANG_GENERAL)
        access.Quit()
    except:
        print("Did not able to create Database File, Please install Access Database Driver and try again!!")

    try:
        conn_str = ('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' f'DBQ={pathOfFolderOutput}{srNumber}_Data_OutPut_{current_date}/DataCompare_Database.mdb;')
        cnxn = pyodbc.connect(conn_str)
        print('Connection Created with .mdb file...')
    except:
        print('Did not find DataCompare_Database.mdb file...')

    print("Process started for Database file, please wait till this screen get close.")

    id = 0
    for file in os.listdir(New_file_path):
        filelocations = New_file_path + '/' + file
        l = len(file)
        if file.endswith(".csv"):
            df= pd.read_csv(filelocations, dtype='object')
            file_name = file[:l - 4]
        elif file.endswith(".xlsx"):
            print("Please keep only 'csv' formatted files in this folder...")
            file_name = file[:l - 5]
        if df.empty:
            new_row = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)
            df = df._append(new_row, ignore_index=True)
        if not df.empty:
            cols1 = df.filter(like='TST', axis=1)
            cols2 = df.filter(like='TST_CMP', axis=1)
            cols3 = df.filter(like='GEN_TST_CMP', axis=1)
            if 'TST' in cols1:
                df["REGION"] = df['TST']
                df.drop("TST", axis=1, inplace=True)
                df = df.rename(columns = {'REGION':'TST'})
            elif 'TST_CMP' in cols2:
                df["REGION"] = df['TST_CMP']
                df.drop("TST_CMP", axis=1, inplace=True)
                df=df.rename(columns = {'REGION':'TST'})
            elif 'GEN_TST_CMP' in cols3:
                df["REGION"] = df['GEN_TST_CMP']
                df.drop("GEN_TST_CMP", axis=1, inplace=True)
                df=df.rename(columns = {'REGION':'TST'})

            if file_name.upper() == 'ACC':
                df["TYPE"] = df["TYPE"].str[0]
                df = df.rename(columns={'DESCRIPTION': 'DESCRIPTION     '})
            if file_name.upper() == 'WGPS':
                df = df.rename(columns={'LNGDSC': 'LNGDSC  '})

            df.fillna("")
            df.replace(np.NaN, '')
            df.replace('nan', '')
            mdb_df = df.copy()
            # mdb file generation
            if file_name.upper() in delete_Cols.keys():
                deleteCol = list(delete_Cols[file_name])
                mdb_df.drop(deleteCol, axis=1, inplace=True)
            if file_name.upper() in file_rename_Mappings1.keys():
                rename_col = list(file_rename_Mappings1[file_name])
                mdb_df.columns = rename_col

            fileName = file_name

            try:
                drop_table_query = f"DROP TABLE {fileName}"
                try:
                    cnxn.execute(drop_table_query)
                except:
                    pass
                mdb_df.rename(columns=lambda x: x.translate(str.maketrans({'-': '_', '#': '_Number', ' ': '_'})), inplace=True)
                create_table_query = f"CREATE TABLE {fileName} ( ID COUNTER PRIMARY KEY, "
                for column_name, dtype in zip(mdb_df.columns, mdb_df.dtypes):
                    max_length = mdb_df[column_name].astype(str).map(len).max()
                    create_table_query += f"  {column_name} VARCHAR ({max_length}),"
                create_table_query = create_table_query.rstrip(",") + ")"
                cnxn.execute(create_table_query)

                mdb_df.fillna('', inplace=True)

                for row in mdb_df.itertuples(index=False):
                    id += 1
                    insert_query = f"INSERT INTO {fileName} VALUES ({id},{','.join(['?' for _ in mdb_df.columns])})"
                    cnxn.execute(insert_query, row)
                cnxn.commit()
            except:
                pass
            #mbd file generation completed
            if not df.empty:
                max_lengths = df.fillna('').astype(str).apply(lambda x: x.str.len()).max().astype(int)
                for col in df.columns:
                    df[col] = df[col].fillna('').astype(str).str.ljust(max_lengths[col])
            df.to_csv(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}/{srNumber}_Text Files/{file_name}.txt', index=False, sep='\t',na_rep='')
    

    #  govind kushwahaclosing cursor
    try:
        cnxn.close()
    except:
        pass
    print(f'DataCompare_Database.mdb and Text file generated, please check it')
    print(f'Process started for generating {srNumber}_Data Compare file..')


    # govind kushwaha XXT_Data Compare File Generating here...

    df = pd.DataFrame()
    HowToReadThisReport = pd.DataFrame()

    heading = 'How to read the Data Compare Report.'
    read = ''' 
    Each tab represents an analysis item. When printing this spreadsheet, any one tab may span multiple pages.

    On each 'tab' or 'page', there is a column with the heading 'Total Of ID'.

    To the left of this column there will be two or three columns of data. The tab label or page heading will define the analysis item. (i.e. Hours & Earnings, Memos, Deductions, etc.)

    To the right of the 'Total Of ID' column there is a column for each ADP company code that is being compared to the columns on the left.
    Under each company code will be the number 1 or a blank. If there is a number 1, that company code's setup matches the colums to the left exactly, on that line. If the box under the company code is blank, that company does not have that item setup.

    The 'Total Of ID' column displays a number on each line. The number on each line is the total count of all company codes on that line that satisfy the analysis item to the left.
    '''

    HowToReadThisReport._set_value(2, heading, read)
    HowToReadThisReport = HowToReadThisReport.style.set_properties(**{'text-wrap': 'normal', 'white-space': 'pre-wrap'})


    fileSheetName = {
    'ACC'	: 'Accumulators',
    'DEDL'	: 'Deductions',
    'DSCH'	: 'Deduction Schedule',
    'HECL'	: 'Hours and Earning',
    'MEML'	: 'Memos',
    'GRID'	: 'Tax Grids',
    'SCBA'	: 'Special Cals Benefit Accruals',
    'SCADJ'	: 'Special Cals Adjustments',
    'SMHD'  : 'Summary Headers',
    'LOCID' : 'Local Tax Codes',
    'WGPS'  : 'Wage Garnishment',
    'CUST'  : 'Customs',
    'INET'  : 'Internet Access',
    'MT'    : 'Meals & Tips',
    'PAYEC' : 'Pay Statement',
    'PCAL'  : 'Payroll Schedule',
    'SSCH'	: 'Special Effects',
    'CAFE'	: 'Cafe Plans',
    'SPC'	: 'Special Compensation',
    'BNK'	: 'Banking Information',
    'SCSUM'	: 'Special Calc Summary',
    'SCFM'	: 'Special Calc Formulas',
    'SCQU'	: 'Special Calc Qualifiers',
    'AT'	: 'Allowed & Taken',
    'GTL'	: 'Group Term Life',
    'GEN'	: 'General Information',
    'GRS'	: 'Gross Calc Info',
    'TLM'	: 'Time & Labor Mngt',
    'STUBL'	: 'Pay Stub Liberals',
    'MISC'	: 'Miscellanous',
    'MRP'	: 'Miscellaneous Reports',
    'RSCH'	: 'Report Schedule',
    'TRCO'	: 'Tax Report Comp Info',
    'CPCDA' : 'Custom Control Cards'

    }

    file_rename_Mappings = {'TRCO': ['QAL-PN','AC-STS','AC-CD','ACCUM DESCRIPTION','CLR','TAXOPT','FLD-SYM','Total Of ID'],
                            'AAC': ['CD', 'DESCRIPTION','CLR','TAXOPT','LDGR','INCEXC','SYMB','AMC','EXP','DWN','Total Of ID'],
    'SCSUM' : ['RES', 'MTH-NUM', 'MTH-DSC', 'CALC-STS', 'ALT-RES', 'CLC-LVL',	'Total Of ID'],
    'SCFM' : ['RES', 'MTH', 'LIN#', 'OPR-A', 'FLD-SYM-A', 'OPR-B', 'FLD-SYM-B', 'OPR-C', 'FLD-SYM-C', 'INT-RES',	'Total Of ID'],
    'SCADJ' : ['RSLT-SYMB-C', 'MTH-NUM', 'FLD-OR-RES', 'TYP', 'AMT', 'ACC-FLD-SYMB-C',	'Total Of ID'],
    'SCQU' : ['RES', 'MTH', 'QL1', 'FR1', 'TO1', 'CD1', 'VAL1', 'QL2', 'FR2', 'TO2', 'CD2', 'VAL2', 'QL3', 'FR3', 'TO3', 'CD3', 'VAL3', 'Total Of ID'],
    'GTL' : ['GTLTYPE', 'COD', 'TAXOPT', 'CLCOCR', 'STDHRS', 'INSFCT', 'ANNMAX', 'RNDNG', 'BDAY', 'EEDED', 'Total Of ID'],
    'SCBA' : ['RES', 'MTH', 'TRRST', 'CURACR', 'CD', 'LIT', 'TYP', 'BNTYP', 'RDTRDOL', 'TRN-TIM', 'TRICR', 'TRAMT', 'EXCTKN', 'UNALLAMT',	'Total Of ID'],
    'SMHD' : ['DPT-NUM', 'DPT-DSC',	'Total Of ID'],
    'LOCID': ['STATECD', 'LOCALTXCD', 'Expr1', 'LOCALNAM',	'Total Of ID'],
    'GEN' : ['VER', 'DEPTSIZ', 'QTR', 'CSR', 'ST-RCP', 'CTY-RCP', 'MJUR', 'PCS',	'Total Of ID'],
    'GRS' : ['REDSTHR1', 'REDSTHR2', 'REDSTHR3', 'SALOTHRS', 'CLCRT2', 'GRSCLC2', 'GRSCLC3', 'GRSCLC4', 'ALLUSE', 'STDHRS',	'Total Of ID'],
    'TLM' : ['TOTALTIME', 'EPIP',	'OTH_TLM', 'Total Of ID'],
    'STUBL' : ['FLD-SYM', 'LNG-DSC-COI-L1', 'LNG-DSC-COI-L2',	'Total Of ID'],
    'PAYEC' : ['OMNI#', 'PRINTAUDMSG', 'CHKSIGNING', 'STUFFSEAL',	'Total Of ID'],
    'CUST' : ['CUSTNETPROC', 'CUSTPROG', 'CUSTREPORTS', 'NOSCUSTIND', 'CUSTCHECK',	'Total Of ID'],
    'CPCDA' : ['CUSTOM', 'REPORT', 'CTRL', 'REC', 'SEQ#', 'CTRLCARDS',	'Total Of ID'],
    'MRP' : ['AUDRPT', 'WRKSHTSRT', 'MAJSRT', 'INTSRT', 'SRTBY', 'ACTFLGSEQ', 'TAXSVCRPT', 'NHIRECOMPL', 'HIREDT', 'BRTHDT', 'LOADT',	'Total Of ID'],
    'RSCH' : ['RSG', 'RUN-FRQ', 'DET-SAV', 'DELIVERY', 'FOR-PER-COD', 'WK-DAY-COD', 'STA-COD',	'Total Of ID'],
    'PCAL' : ['PAYFRQ', 'WK', 'PEDAT', 'INDAT', 'OUTDAT', 'PAYDAT',	'Total Of ID'],
    'INET' : ['IPAY', 'IREPORTS', 'SUP PAPER', 'PR QUICKVIEW', 'REMCTRL', 'SELF-SVC', 'W2_REISSUE', 'P_FREEVERS', 'P_CONTMAN', 'ORG_OID', 'RES_ADPREG', 'REALTIME',	'Total Of ID']
                            }


    for file in os.listdir(New_file_path):
        file_name = file
        if (file_name.upper() == 'LIST.XLSX' or file_name.upper() == 'LIST.CSV') and (file.endswith(".xlsx") or file.endswith(".csv")):
            filelocations = New_file_path + '/' + file
            try:
                try:
                    df_company = pd.read_csv(filelocations, dtype='object')
                except:
                    df_company = pd.read_csv(filelocations, dtype='object', encoding='latin')
            except:
                df_company = pd.read_excel(filelocations, dtype='object')
            df_company["REGION"] = df_company['TST']
            df_company.drop("TST", axis=1, inplace=True)
            df_company = df_company.rename(columns={'REGION': 'TST'})
            df_company.drop_duplicates(keep='first', inplace=True)

            compCode_df = df_company["CMP"].tolist()
            Temp_compCode = compCode_df


    if df_company.empty:
        print("please keep Company code list file with name 'List.csv' in this folder and try again..")
        exit()
    writer = pd.ExcelWriter(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}/{srNumber}_Data Compare.xlsx', engine='xlsxwriter')
    HowToReadThisReport.to_excel(writer, sheet_name='How To Read Report', index=False)
    try:
        df_company = df_company.style.set_properties(**{'font-family': 'Arial, Helvetica, sans-serif', 'font-size': '12px'})
    except:
        pass
    df_company.to_excel(writer, sheet_name='Company Code List', index=False)

    for file in os.listdir(New_file_path):
        compCode_df = Temp_compCode
        file_name = file
        filelocations1 = New_file_path + '/' + file
        if (file_name.upper() != 'LIST.XLSX' or file_name.upper() != 'LIST.CSV' or file_name != f'{srNumber}_Data Compare.xlsx') and file_name.endswith('.csv'):
            try:
                try:
                    df_Accumulator = pd.read_csv(filelocations1, dtype='object')
                    fileNamesForSheets = file_name[:-4]
                except:
                    df_Accumulator = pd.read_csv(filelocations1, dtype='object', encoding='latin')
                    fileNamesForSheets = file_name[:-4]
            except:
                df_Accumulator = pd.read_excel(filelocations1, dtype='object')
                fileNamesForSheets = file_name[:-5]

            if df_Accumulator.empty:
                s = pd.Series(None, index=df_Accumulator.columns)
                # Appending empty series to df
                df_Accumulator = df_Accumulator._append(s, ignore_index=True)

            if not df_Accumulator.empty and fileNamesForSheets.strip().upper() in fileSheetName.keys():

                if fileNamesForSheets.upper() in delete_Cols.keys():
                    deleteCol = list(delete_Cols[fileNamesForSheets])
                    df_Accumulator.drop(deleteCol, axis=1, inplace=True)


                cols1 = df_Accumulator.filter(like='TST', axis=1)
                cols2 = df_Accumulator.filter(like='TST_CMP', axis=1)
                cols3 = df_Accumulator.filter(like='GEN_TST_CMP', axis=1)
                if 'TST' in cols1:
                    df_Accumulator.drop(["TST", "REGION"], axis=1, inplace=True)
                elif 'TST_CMP' in cols2:
                    df_Accumulator.drop(["TST_CMP", "REGION"], axis=1, inplace=True)
                elif 'GEN_TST_CMP' in cols3:
                    df_Accumulator.drop(["GEN_TST_CMP", "REGION"], axis=1, inplace=True)

                cmpCol1 = df_Accumulator.filter(like='CMP', axis=1)
                cmpCol2 = df_Accumulator.filter(like='COCODE', axis=1)
                cmpCol3 = df_Accumulator.filter(like='CO_CODE', axis=1)

                if 'CMP' in cmpCol1:
                    if file_name.upper() == 'ACC.XLSX' or file_name.upper() == 'ACC.CSV':
                        df_Accumulator.drop(["STS", 'TYPE'], axis=1, inplace=True)
                    acc_df = df_Accumulator
                    acc_df = acc_df.drop_duplicates(keep='last')

                    df_Accumulator.drop("CMP", axis=1, inplace=True)
                    totalInsertIndex = len(acc_df.columns)
                    for code in compCode_df:
                        if code in acc_df.columns:
                            acc_df.rename(columns={code: f'{code}_1'}, inplace=True)
                            df_Accumulator.rename(columns={code: f'{code}_1'}, inplace=True)
                        acc_df[code] = acc_df['CMP'] == code

                elif 'COCODE' in cmpCol2:
                    acc_df = df_Accumulator
                    acc_df = acc_df.drop_duplicates(keep='last')

                    df_Accumulator.drop("COCODE", axis=1, inplace=True)
                    totalInsertIndex = len(acc_df.columns)
                    for code in compCode_df:
                        if code in acc_df.columns:
                            acc_df.rename(columns={code: f'{code}_1'}, inplace=True)
                            df_Accumulator.rename(columns={code: f'{code}_1'}, inplace=True)
                        acc_df[code] = acc_df['COCODE'] == code

                elif 'CO_CODE' in cmpCol2:
                    acc_df = df_Accumulator
                    acc_df = acc_df.drop_duplicates(keep='last')

                    df_Accumulator.drop("CO_CODE", axis=1, inplace=True)
                    totalInsertIndex = len(acc_df.columns)
                    for code in compCode_df:
                        if code in acc_df.columns:
                            acc_df.rename(columns={code: f'{code}_1'}, inplace=True)
                            df_Accumulator.rename(columns={code: f'{code}_1'}, inplace=True)
                        acc_df[code] = acc_df['CO_CODE'] == code

                acc_df = acc_df.drop_duplicates(keep='last')
                groupCol = df_Accumulator.columns.tolist()
                groupCol = list(groupCol)
                acc_df.fillna(' ', inplace=True)

                grouped_df = acc_df.groupby(groupCol, as_index=False).sum()

                grouped_df.replace(True, '1', inplace=True)
                grouped_df.replace(False, '0', inplace=True)

                df['Total Of ID'] = grouped_df[compCode_df].astype(int).sum(axis=1)
                grouped_df.insert(totalInsertIndex, 'Total Of ID', df['Total Of ID'])

                cmpCol1 = grouped_df.filter(like='CMP', axis=1)
                cmpCol2 = grouped_df.filter(like='COCODE', axis=1)
                cmpCol3 = df_Accumulator.filter(like='CO_CODE', axis=1)

                if 'CMP' in cmpCol1:
                    grouped_df.drop('CMP', axis=1, inplace=True)

                elif 'COCODE' in cmpCol2:
                    grouped_df.drop('COCODE', axis=1, inplace=True)

                elif 'CO_CODE' in cmpCol2:
                    grouped_df.drop('CO_CODE', axis=1, inplace=True)

                if fileNamesForSheets.upper() in file_rename_Mappings.keys():
                    rename_col = list(file_rename_Mappings[fileNamesForSheets])
                    compCode_df = [str(element) + ' ' if element in rename_col else element for element in compCode_df]
                    rename_col = rename_col + compCode_df
                    try:
                        grouped_df.columns = rename_col
                    except:
                        print(f'Please check this file: {file_name}')

                try:
                    grouped_df[compCode_df] = grouped_df[compCode_df].replace(0, '')
                except:
                    if fileNamesForSheets.upper() in fileSheetName.keys():
                        sheetName = str(fileSheetName[fileNamesForSheets])
                        print('Please remove the zeros from sheet- ', sheetName)
    
                fileNamesForSheets = fileNamesForSheets.upper()
                if fileNamesForSheets.upper() in fileSheetName.keys():
                    sheetName = str(fileSheetName[fileNamesForSheets])
                try:
                    grouped_df = grouped_df.style.set_properties(**{'font-family': 'Arial, Helvetica, sans-serif', 'font-size': '12px'})
                except:
                    if fileNamesForSheets.upper() in fileSheetName.keys():
                        sheetName = str(fileSheetName[fileNamesForSheets])
                        print('Please format sheet- ', sheetName)
                grouped_df.to_excel(writer, sheet_name=f'{sheetName}', index=False)

    writer._save()

    print(f'Process completed, please check files in {srNumber}_Data OutPut, Enjoy your Day!!')

    tk.messagebox.showwarning('Process Completed!!', f'You can find your Reports in {pathOfFolderOutput}')


#govind kushwaha Zip Fie Only

elif var == 2:
    access = win32com.client.Dispatch('Access.Application')
    from comtypes.gen import Access

    def  open_csv_file():
        global file_path
        file_pat = filedialog.askopenfilename()
        file_path.set(file_pat)

    def submit_form():
        global clientNameOrId, file_path
        clientNameOrId = enter_datalake_user_id.get()
        file_path = file_path_entery.get()
        clientNameOrId = clientNameOrId
        file_path = file_path
        root.destroy()
    root = tk.Tk()
    root.title('Fill the Details!')
    label_datalake_user_id = tk.Label(root, text='Client ID/Name: ', font=("Helvetica", 14))
    label_datalake_user_id.grid(row=0, column=0)

    enter_datalake_user_id = tk.Entry(root, font=("Helvetica", 14))
    enter_datalake_user_id.grid(row=0, column=1)

    file_path = tk.StringVar()
    file_path_lable = tk.Label(root,text='Select eCCI Zip File:', font=("Helvetica", 14))
    file_path_lable.grid(row=1, column=0)
    file_path_entery = tk.Entry(root, textvariable=file_path, state="readonly", font=("Helvetica", 14))
    file_path_entery.grid(row=1, column=1)
    file_button = tk.Button(root, text='Browse File', command= open_csv_file, font=("Helvetica", 14))
    file_button.grid(row=1, column=2, pady=10,padx=20)
    submit_button = tk.Button(root, text="Submit", command=submit_form, font=("Helvetica", 14))
    submit_button.grid(row=2, columnspan=4, pady=10, padx=20)

    for i in range(2):
        root.grid_rowconfigure(i, weight=2)
    for i in range(2):
        root.grid_columnconfigure(i, weight=2)
    window_width = 600
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    root.mainloop()

    if clientNameOrId == '':
        srNumber = "Client_Name"
    else:
        srNumber = clientNameOrId
    current_dates = str(datetime.datetime.now().strftime('%Y%m%d'))

    if file_path == '':
        print('Please select Zip File..')
        exit()
    zip_file_name = os.path.basename(file_path)
    pathOfFolder = file_path[:-len(zip_file_name)]
    pathOfFolderOutput = pathOfFolder

    mdb_df = pd.DataFrame()


    delete_Cols = {
    'SCSUM' : ['RESDESC'],
    'SCFM' : ['PRIORITY'],
    'SCADJ' : ['LN', 'STEP'],
    'SCQU' : ['PRIORITY'],
    'LOCID' : ['PA32_RLO'],
    'MRP': ['TERM_DT'],
    'GRS' : ['AUTO_PAY', 'BY_PAY_GRP', 'FLD3', 'DED_FLD3', 'FLD4', 'DED_FLD4', 'FLD5', 'DED_FLD5', 'REDUCE_HR4', 'REDUCE_HR5', 'REDUCE_HR6', 'REDUCE_HR7', 'REDUCE_HR8', 'REDUCE_HR9', 'REDUCE_HR10', 'REDUCE_HR11', 'REDUCE_HR12', 'REDUCE_HR13', 'REDUCE_HR14', 'REDUCE_HR15', 'REDUCE_HR16', 'REDUCE_HR17', 'REDUCE_HR18', 'REDUCE_HR19', 'REDUCE_HR20', 'REDUCE_HR21']
    }

    file_rename_Mappings1 = {'TRCO': ['TST','CMP','QAL-PN','AC-STS','AC-CD','ACCUM DESCRIPTION','CLR','TAXOPT','FLD-SYM'],
                            'AAC': ['TST','CMP','CD', 'DESCRIPTION','CLR','TAXOPT','LDGR','INCEXC','SYMB','AMC','EXP','DWN'],
    'SCSUM' : ['TST','CMP','RES', 'MTH-NUM', 'MTH-DSC', 'CALC-STS', 'ALT-RES', 'CLC-LVL'],
    'SCFM' : ['TST','CMP','RES', 'MTH', 'LIN#', 'OPR-A', 'FLD-SYM-A', 'OPR-B', 'FLD-SYM-B', 'OPR-C', 'FLD-SYM-C', 'INT-RES'],
    'SCADJ' : ['TST','CMP','RSLT-SYMB-C', 'MTH-NUM', 'FLD-OR-RES', 'TYP', 'AMT', 'ACC-FLD-SYMB-C'],
    'SCQU' : ['TST','CMP','RES', 'MTH', 'QL1', 'FR1', 'TO1', 'CD1', 'VAL1', 'QL2', 'FR2', 'TO2', 'CD2', 'VAL2', 'QL3', 'FR3', 'TO3', 'CD3', 'VAL3'],
    'GTL' : ['TST','CMP','GTLTYPE', 'COD', 'TAXOPT', 'CLCOCR', 'STDHRS', 'INSFCT', 'ANNMAX', 'RNDNG', 'BDAY', 'EEDED'],
    'SCBA' : ['TST','CMP','RES', 'MTH', 'TRRST', 'CURACR', 'CD', 'LIT', 'TYP', 'BNTYP', 'RDTRDOL', 'TRN-TIM', 'TRICR', 'TRAMT', 'EXCTKN', 'UNALLAMT'],
    'SMHD' : ['TST','CMP','DPT-NUM', 'DPT-DSC'],
    'LOCID': ['TST','CMP','STATECD', 'LOCALTXCD', 'Expr1', 'LOCALNAM'],
    'GEN' : ['TST','CMP','VER', 'DEPTSIZ', 'QTR', 'CSR', 'ST-RCP', 'CTY-RCP', 'MJUR', 'PCS'],
    'GRS' : ['TST','CMP','REDSTHR1', 'REDSTHR2', 'REDSTHR3', 'SALOTHRS', 'CLCRT2', 'GRSCLC2', 'GRSCLC3', 'GRSCLC4', 'ALLUSE', 'STDHRS'],
    'TLM' : ['TST','CMP','TOTALTIME', 'EPIP',	'OTH_TLM'],
    'STUBL' : ['TST','CMP','FLD-SYM', 'LNG-DSC-COI-L1', 'LNG-DSC-COI-L2'],
    'PAYEC' : ['TST','CMP','OMNI#', 'PRINTAUDMSG', 'CHKSIGNING', 'STUFFSEAL'],
    'CUST' : ['TST','CMP','CUSTNETPROC', 'CUSTPROG', 'CUSTREPORTS', 'NOSCUSTIND', 'CUSTCHECK'],
    'CPCDA' : ['TST','CMP','CUSTOM', 'REPORT', 'CTRL', 'REC', 'SEQ#', 'CTRLCARDS'],
    'MRP' : ['TST','CMP','AUDRPT', 'WRKSHTSRT', 'MAJSRT', 'INTSRT', 'SRTBY', 'ACTFLGSEQ', 'TAXSVCRPT', 'NHIRECOMPL', 'HIREDT', 'BRTHDT', 'LOADT'],
    'RSCH' : ['TST','CMP','RSG', 'RUN-FRQ', 'DET-SAV', 'DELIVERY', 'FOR-PER-COD', 'WK-DAY-COD', 'STA-COD'],
    'PCAL' : ['TST','CMP','PAYFRQ', 'WK', 'PEDAT', 'INDAT', 'OUTDAT', 'PAYDAT'],
    'INET' : ['TST','CMP','IPAY', 'IREPORTS', 'SUP PAPER', 'PR QUICKVIEW', 'REMCTRL', 'SELF-SVC', 'W2_REISSUE', 'P_FREEVERS', 'P_CONTMAN', 'ORG_OID', 'RES_ADPREG', 'REALTIME']
                            }

    current_date = str(datetime.datetime.now().strftime('%d-%m-%Y'))


    if not os.path.exists(f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}'):
        os.makedirs(f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}')
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}')

    if os.path.exists(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}'):
        try:
            os.rmdir(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}')
        except OSError as e:
            print(f'Folder {srNumber}_Data_OutPut does not exist, creating new folder....')

    if not os.path.exists(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}'):
        os.makedirs(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}')
        os.makedirs(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}/{srNumber}_Text Files')

    New_file_path = f'{pathOfFolderOutput}/{zip_file_name[:-4]}_{current_date}'
    try:
        DB_Engine = access.DBEngine
        DB_Engine.CreateDatabase(f'{pathOfFolderOutput}{srNumber}_Data_OutPut_{current_date}/DataCompare_Database.mdb', Access.DB_LANG_GENERAL)
        access.Quit()
    except:
        print("Not able to Create Access Database file!!, Please install Access Database Driver and Try again.")
    # govind kushwaha current_directory = os.getcwd()

    try:
        conn_str = ('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' f'DBQ={pathOfFolderOutput}{srNumber}_Data_OutPut_{current_date}/DataCompare_Database.mdb;')
        cnxn = pyodbc.connect(conn_str)
        print('Connection Created with .mdb file...')
    except:
        print('Did not find DataCompare_Database.mdb file...')

    print("Process started for database file, please wait till this screen get close.")

    id = 0
    for file in os.listdir(New_file_path):
        filelocations = New_file_path + '/' + file
        l = len(file)
        if file.endswith(".csv"):
            df= pd.read_csv(filelocations, dtype='object')
            file_name = file[:l - 4]
        elif file.endswith(".xlsx"):
            print("Please keep only 'csv' formatted files in this folder...")
            file_name = file[:l - 5]
        if df.empty:
            new_row = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)
            df = df._append(new_row, ignore_index=True)
        if not df.empty:
            cols1 = df.filter(like='TST', axis=1)
            cols2 = df.filter(like='TST_CMP', axis=1)
            cols3 = df.filter(like='GEN_TST_CMP', axis=1)
            if 'TST' in cols1:
                df["REGION"] = df['TST']
                df.drop("TST", axis=1, inplace=True)
                df = df.rename(columns = {'REGION':'TST'})
            elif 'TST_CMP' in cols2:
                df["REGION"] = df['TST_CMP']
                df.drop("TST_CMP", axis=1, inplace=True)
                df=df.rename(columns = {'REGION':'TST'})
            elif 'GEN_TST_CMP' in cols3:
                df["REGION"] = df['GEN_TST_CMP']
                df.drop("GEN_TST_CMP", axis=1, inplace=True)
                df=df.rename(columns = {'REGION':'TST'})

            if file_name.upper() == 'ACC':
                df["TYPE"] = df["TYPE"].str[0]
                df = df.rename(columns={'DESCRIPTION': 'DESCRIPTION     '})
            if file_name.upper() == 'WGPS':
                df = df.rename(columns={'LNGDSC': 'LNGDSC  '})

            df.fillna("", inplace=True)
            df.replace('nan', '')
            mdb_df = df.copy()
            # govind kushwaha mdb file generation
            if file_name.upper() in delete_Cols.keys():
                deleteCol = list(delete_Cols[file_name])
                mdb_df.drop(deleteCol, axis=1, inplace=True)
            if file_name.upper() in file_rename_Mappings1.keys():
                rename_col = list(file_rename_Mappings1[file_name])
                mdb_df.columns = rename_col

            fileName = file_name

            try:
                drop_table_query = f"DROP TABLE {fileName}"
                try:
                    cnxn.execute(drop_table_query)
                except:
                    pass
                mdb_df.rename(columns=lambda x: x.translate(str.maketrans({'-': '_', '#': '_Number', ' ': '_'})), inplace=True)
                create_table_query = f"CREATE TABLE {fileName} ( ID COUNTER PRIMARY KEY, "
                for column_name, dtype in zip(mdb_df.columns, mdb_df.dtypes):
                    max_length = mdb_df[column_name].astype(str).map(len).max()
                    create_table_query += f"  {column_name} VARCHAR ({max_length}),"
                create_table_query = create_table_query.rstrip(",") + ")"
                cnxn.execute(create_table_query)

                mdb_df.fillna('', inplace=True)

                for row in mdb_df.itertuples(index=False):
                    id += 1
                    insert_query = f"INSERT INTO {fileName} VALUES ({id},{','.join(['?' for _ in mdb_df.columns])})"
                    cnxn.execute(insert_query, row)

                cnxn.commit()
            except:
                print("Got error while inserting data in Access Database file.")
            #mbd file generation completed
            if not df.empty:
                max_lengths = df.fillna('').astype(str).apply(lambda x: x.str.len()).max().astype(int)
                for col in df.columns:
                    df[col] = df[col].fillna('').astype(str).str.ljust(max_lengths[col])
            df.to_csv(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}/{srNumber}_Text Files/{file_name}.txt', index=False, sep='\t',na_rep='')
    # govind kushwaha closing cursor
    try:
        cnxn.close()
        print("..")
    except:
        pass
    print(f'DataCompare_Database.mdb and Text file generated!!')

    print(f'Process started for generating {srNumber}_Data Compare file..')


    # govind kushwaha Data Compare File Generating here...

    df = pd.DataFrame()
    HowToReadThisReport = pd.DataFrame()

    heading = 'How to read the Data Compare Report.'
    read = ''' 
    Each tab represents an analysis item. When printing this spreadsheet, any one tab may span multiple pages.

    On each 'tab' or 'page', there is a column with the heading 'Total Of ID'.

    To the left of this column there will be two or three columns of data. The tab label or page heading will define the analysis item. (i.e. Hours & Earnings, Memos, Deductions, etc.)

    To the right of the 'Total Of ID' column there is a column for each ADP company code that is being compared to the columns on the left.
    Under each company code will be the number 1 or a blank. If there is a number 1, that company code's setup matches the colums to the left exactly, on that line. If the box under the company code is blank, that company does not have that item setup.

    The 'Total Of ID' column displays a number on each line. The number on each line is the total count of all company codes on that line that satisfy the analysis item to the left.
    '''

    HowToReadThisReport._set_value(2, heading, read)
    HowToReadThisReport = HowToReadThisReport.style.set_properties(**{'text-wrap': 'normal', 'white-space': 'pre-wrap'})


    fileSheetName = {
    'ACC'	: 'Accumulators',
    'DEDL'	: 'Deductions',
    'DSCH'	: 'Deduction Schedule',
    'HECL'	: 'Hours and Earning',
    'MEML'	: 'Memos',
    'GRID'	: 'Tax Grids',
    'SCBA'	: 'Special Cals Benefit Accruals',
    'SCADJ'	: 'Special Cals Adjustments',
    'SMHD'  : 'Summary Headers',
    'LOCID' : 'Local Tax Codes',
    'WGPS'  : 'Wage Garnishment',
    'CUST'  : 'Customs',
    'INET'  : 'Internet Access',
    'MT'    : 'Meals & Tips',
    'PAYEC' : 'Pay Statement',
    'PCAL'  : 'Payroll Schedule',
    'SSCH'	: 'Special Effects',
    'CAFE'	: 'Cafe Plans',
    'SPC'	: 'Special Compensation',
    'BNK'	: 'Banking Information',
    'SCSUM'	: 'Special Calc Summary',
    'SCFM'	: 'Special Calc Formulas',
    'SCQU'	: 'Special Calc Qualifiers',
    'AT'	: 'Allowed & Taken',
    'GTL'	: 'Group Term Life',
    'GEN'	: 'General Information',
    'GRS'	: 'Gross Calc Info',
    'TLM'	: 'Time & Labor Mngt',
    'STUBL'	: 'Pay Stub Liberals',
    'MISC'	: 'Miscellanous',
    'MRP'	: 'Miscellaneous Reports',
    'RSCH'	: 'Report Schedule',
    'TRCO'	: 'Tax Report Comp Info',
    'CPCDA' : 'Custom Control Cards'

    }

    file_rename_Mappings = {'TRCO': ['QAL-PN','AC-STS','AC-CD','ACCUM DESCRIPTION','CLR','TAXOPT','FLD-SYM','Total Of ID'],
                            'AAC': ['CD', 'DESCRIPTION','CLR','TAXOPT','LDGR','INCEXC','SYMB','AMC','EXP','DWN','Total Of ID'],
    'SCSUM' : ['RES', 'MTH-NUM', 'MTH-DSC', 'CALC-STS', 'ALT-RES', 'CLC-LVL',	'Total Of ID'],
    'SCFM' : ['RES', 'MTH', 'LIN#', 'OPR-A', 'FLD-SYM-A', 'OPR-B', 'FLD-SYM-B', 'OPR-C', 'FLD-SYM-C', 'INT-RES',	'Total Of ID'],
    'SCADJ' : ['RSLT-SYMB-C', 'MTH-NUM', 'FLD-OR-RES', 'TYP', 'AMT', 'ACC-FLD-SYMB-C',	'Total Of ID'],
    'SCQU' : ['RES', 'MTH', 'QL1', 'FR1', 'TO1', 'CD1', 'VAL1', 'QL2', 'FR2', 'TO2', 'CD2', 'VAL2', 'QL3', 'FR3', 'TO3', 'CD3', 'VAL3', 'Total Of ID'],
    'GTL' : ['GTLTYPE', 'COD', 'TAXOPT', 'CLCOCR', 'STDHRS', 'INSFCT', 'ANNMAX', 'RNDNG', 'BDAY', 'EEDED', 'Total Of ID'],
    'SCBA' : ['RES', 'MTH', 'TRRST', 'CURACR', 'CD', 'LIT', 'TYP', 'BNTYP', 'RDTRDOL', 'TRN-TIM', 'TRICR', 'TRAMT', 'EXCTKN', 'UNALLAMT',	'Total Of ID'],
    'SMHD' : ['DPT-NUM', 'DPT-DSC',	'Total Of ID'],
    'LOCID': ['STATECD', 'LOCALTXCD', 'Expr1', 'LOCALNAM',	'Total Of ID'],
    'GEN' : ['VER', 'DEPTSIZ', 'QTR', 'CSR', 'ST-RCP', 'CTY-RCP', 'MJUR', 'PCS',	'Total Of ID'],
    'GRS' : ['REDSTHR1', 'REDSTHR2', 'REDSTHR3', 'SALOTHRS', 'CLCRT2', 'GRSCLC2', 'GRSCLC3', 'GRSCLC4', 'ALLUSE', 'STDHRS',	'Total Of ID'],
    'TLM' : ['TOTALTIME', 'EPIP',	'OTH_TLM', 'Total Of ID'],
    'STUBL' : ['FLD-SYM', 'LNG-DSC-COI-L1', 'LNG-DSC-COI-L2',	'Total Of ID'],
    'PAYEC' : ['OMNI#', 'PRINTAUDMSG', 'CHKSIGNING', 'STUFFSEAL',	'Total Of ID'],
    'CUST' : ['CUSTNETPROC', 'CUSTPROG', 'CUSTREPORTS', 'NOSCUSTIND', 'CUSTCHECK',	'Total Of ID'],
    'CPCDA' : ['CUSTOM', 'REPORT', 'CTRL', 'REC', 'SEQ#', 'CTRLCARDS',	'Total Of ID'],
    'MRP' : ['AUDRPT', 'WRKSHTSRT', 'MAJSRT', 'INTSRT', 'SRTBY', 'ACTFLGSEQ', 'TAXSVCRPT', 'NHIRECOMPL', 'HIREDT', 'BRTHDT', 'LOADT',	'Total Of ID'],
    'RSCH' : ['RSG', 'RUN-FRQ', 'DET-SAV', 'DELIVERY', 'FOR-PER-COD', 'WK-DAY-COD', 'STA-COD',	'Total Of ID'],
    'PCAL' : ['PAYFRQ', 'WK', 'PEDAT', 'INDAT', 'OUTDAT', 'PAYDAT',	'Total Of ID'],
    'INET' : ['IPAY', 'IREPORTS', 'SUP PAPER', 'PR QUICKVIEW', 'REMCTRL', 'SELF-SVC', 'W2_REISSUE', 'P_FREEVERS', 'P_CONTMAN', 'ORG_OID', 'RES_ADPREG', 'REALTIME',	'Total Of ID']
                            }


    for file in os.listdir(New_file_path):
        file_name = file
        if (file_name.upper() == 'LIST.XLSX' or file_name.upper() == 'LIST.CSV') and (file.endswith(".xlsx") or file.endswith(".csv")):
            filelocations = New_file_path + '/' + file
            try:
                try:
                    df_company = pd.read_csv(filelocations, dtype='object')
                except:
                    df_company = pd.read_csv(filelocations, dtype='object', encoding='latin')
            except:
                df_company = pd.read_excel(filelocations, dtype='object')
            df_company["REGION"] = df_company['TST']
            df_company.drop("TST", axis=1, inplace=True)
            df_company = df_company.rename(columns={'REGION': 'TST'})
            df_company.drop_duplicates(keep='first', inplace=True)

            compCode_df = df_company["CMP"].tolist()
            Temp_compCode = compCode_df 

    if df_company.empty:
        print("please keep Company code list file with name 'List.csv' in this folder and try again..")
        exit()
    writer = pd.ExcelWriter(f'{pathOfFolderOutput}/{srNumber}_Data_OutPut_{current_date}/{srNumber}_Data Compare.xlsx', engine='xlsxwriter')
    HowToReadThisReport.to_excel(writer, sheet_name='How To Read Report', index=False)
    try:
        df_company = df_company.style.set_properties(**{'font-family': 'Arial, Helvetica, sans-serif', 'font-size': '12px'})
    except:
        pass
    df_company.to_excel(writer, sheet_name='Company Code List', index=False)

    for file in os.listdir(New_file_path):
        file_name = file
        compCode_df = Temp_compCode
        filelocations1 = New_file_path + '/' + file
        if (file_name.upper() != 'LIST.XLSX' or file_name.upper() != 'LIST.CSV' or file_name != f'{srNumber}_Data Compare.xlsx') and file_name.endswith('.csv'):
            try:
                try:
                    df_Accumulator = pd.read_csv(filelocations1, dtype='object')
                    fileNamesForSheets = file_name[:-4]
                except:
                    df_Accumulator = pd.read_csv(filelocations1, dtype='object', encoding='latin')
                    fileNamesForSheets = file_name[:-4]
            except:
                df_Accumulator = pd.read_excel(filelocations1, dtype='object')
                fileNamesForSheets = file_name[:-5]

            if df_Accumulator.empty:
                s = pd.Series(None, index=df_Accumulator.columns)
                # govind kushwaha Appending empty series to df
                df_Accumulator = df_Accumulator._append(s, ignore_index=True)

            if not df_Accumulator.empty and fileNamesForSheets.strip().upper() in fileSheetName.keys():

                if fileNamesForSheets.upper() in delete_Cols.keys():
                    deleteCol = list(delete_Cols[fileNamesForSheets])
                    df_Accumulator.drop(deleteCol, axis=1, inplace=True)


                cols1 = df_Accumulator.filter(like='TST', axis=1)
                cols2 = df_Accumulator.filter(like='TST_CMP', axis=1)
                cols3 = df_Accumulator.filter(like='GEN_TST_CMP', axis=1)
                if 'TST' in cols1:
                    df_Accumulator.drop(["TST", "REGION"], axis=1, inplace=True)
                elif 'TST_CMP' in cols2:
                    df_Accumulator.drop(["TST_CMP", "REGION"], axis=1, inplace=True)
                elif 'GEN_TST_CMP' in cols3:
                    df_Accumulator.drop(["GEN_TST_CMP", "REGION"], axis=1, inplace=True)

                cmpCol1 = df_Accumulator.filter(like='CMP', axis=1)
                cmpCol2 = df_Accumulator.filter(like='COCODE', axis=1)
                cmpCol3 = df_Accumulator.filter(like='CO_CODE', axis=1)

                if 'CMP' in cmpCol1:
                    if file_name.upper() == 'ACC.XLSX' or file_name.upper() == 'ACC.CSV':
                        df_Accumulator.drop(["STS", 'TYPE'], axis=1, inplace=True)
                    acc_df = df_Accumulator
                    acc_df = acc_df.drop_duplicates(keep='last')

                    df_Accumulator.drop("CMP", axis=1, inplace=True)
                    totalInsertIndex = len(acc_df.columns)
                    for code in compCode_df:
                        if code in acc_df.columns:
                            acc_df.rename(columns={code: f'{code}_1'}, inplace=True)
                            df_Accumulator.rename(columns={code: f'{code}_1'}, inplace=True)
                        acc_df[code] = acc_df['CMP'] == code

                elif 'COCODE' in cmpCol2:
                    acc_df = df_Accumulator
                    acc_df = acc_df.drop_duplicates(keep='last')

                    df_Accumulator.drop("COCODE", axis=1, inplace=True)
                    totalInsertIndex = len(acc_df.columns)
                    for code in compCode_df:
                        if code in acc_df.columns:
                            acc_df.rename(columns={code: f'{code}_1'}, inplace=True)
                            df_Accumulator.rename(columns={code: f'{code}_1'}, inplace=True)
                        
                        acc_df[code] = acc_df['COCODE'] == code

                elif 'CO_CODE' in cmpCol2:
                    acc_df = df_Accumulator
                    acc_df = acc_df.drop_duplicates(keep='last')

                    df_Accumulator.drop("CO_CODE", axis=1, inplace=True)
                    totalInsertIndex = len(acc_df.columns)
                    for code in compCode_df:
                        if code in acc_df.columns:
                            acc_df.rename(columns={code: f'{code}_1'}, inplace=True)
                            df_Accumulator.rename(columns={code: f'{code}_1'}, inplace=True)
                        
                        acc_df[code] = acc_df['CO_CODE'] == code

                acc_df = acc_df.drop_duplicates(keep='last')
                groupCol = df_Accumulator.columns.tolist()
                groupCol = list(groupCol)
                acc_df.fillna(' ', inplace=True)

                grouped_df = acc_df.groupby(groupCol, as_index=False).sum()

                grouped_df.replace(True, '1', inplace=True)
                grouped_df.replace(False, '0', inplace=True)

                df['Total Of ID'] = grouped_df[compCode_df].astype(int).sum(axis=1)
                grouped_df.insert(totalInsertIndex, 'Total Of ID', df['Total Of ID'])

                cmpCol1 = grouped_df.filter(like='CMP', axis=1)
                cmpCol2 = grouped_df.filter(like='COCODE', axis=1)
                cmpCol3 = df_Accumulator.filter(like='CO_CODE', axis=1)

                if 'CMP' in cmpCol1:
                    grouped_df.drop('CMP', axis=1, inplace=True)

                elif 'COCODE' in cmpCol2:
                    grouped_df.drop('COCODE', axis=1, inplace=True)

                elif 'CO_CODE' in cmpCol2:
                    grouped_df.drop('CO_CODE', axis=1, inplace=True)

                if fileNamesForSheets.upper() in file_rename_Mappings.keys():
                    rename_col = list(file_rename_Mappings[fileNamesForSheets])
                    
                    compCode_df = [str(element) + ' ' if element in rename_col else element for element in compCode_df]
                    rename_col = rename_col + compCode_df
                    
                    try:
                        grouped_df.columns = rename_col
                    except:
                        print(f'Please check this file columns: {file_name}')
                try:
                    grouped_df[compCode_df] = grouped_df[compCode_df].replace(0, '')
                except:
                    if fileNamesForSheets.upper() in fileSheetName.keys():
                        sheetName = str(fileSheetName[fileNamesForSheets])
                        print('Please remove the zeros from sheet- ', sheetName)
                
                fileNamesForSheets = fileNamesForSheets.upper()
                if fileNamesForSheets.upper() in fileSheetName.keys():
                    sheetName = str(fileSheetName[fileNamesForSheets])
                try:
                    grouped_df = grouped_df.style.set_properties(**{'font-family': 'Arial, Helvetica, sans-serif', 'font-size': '12px'})
                except:
                     if fileNamesForSheets.upper() in fileSheetName.keys():
                        sheetName = str(fileSheetName[fileNamesForSheets])
                        print('Please format sheet- ', sheetName)
                
                grouped_df.to_excel(writer, sheet_name=f'{sheetName}', index=False)

    writer._save()

    print(f'Process completed, please check files in {srNumber}_Data OutPut, Enjoy your Day!!')

    tk.messagebox.showwarning('Process Completed!!', f'You can find your Reports in {pathOfFolderOutput[:-1]}')

# govind kushwaha
