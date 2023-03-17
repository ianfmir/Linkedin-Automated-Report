import pyautogui
import pytesseract
import time
import pandas as pd
from datetime import datetime

# The browser should be on full screen to capture the information on the screenshot, return the text from the webpage
def get_txt_from_screenshot(link):
    pyautogui.press('win')
    time.sleep(0.5)
    pyautogui.write('google chrome')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.write(link)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.1)
    pyautogui.press('x')
    time.sleep(5)
    im1 = pyautogui.screenshot(region=(x,y,w,h))
    time.sleep(1)
    pyautogui.hotkey('alt','f4')
    txt = pytesseract.image_to_string(im1,lang="eng")
    return txt

def get_company_name(txt):
    company_name = ''
    index_line_break = txt.find("\n")
    for i in range(index_line_break):
        company_name += txt[i]
    return company_name

def get_number_of_followers(txt):
    string_before_followers, string_after_followers = txt.split("followers",1)
    split_string_followers = string_before_followers.split()
    return split_string_followers[-1]

def get_number_of_employees(txt):
    string_before_employees, string_after_employees = txt.split("employees",1)
    split_string_employees = string_before_employees.split()
    return split_string_employees[-1]

# Compute limits of the screenshot area
(screen_width, screen_height) = pyautogui.size()
x = screen_width/4.83
y = screen_height/3.3
w = screen_width/2.5
h = screen_height/5

# Open txt file
file = open('companies.txt','r')
file_lines = file.readlines()
number_of_companies = len(file_lines)

# Build the dataframe
dataframe = pd.DataFrame(columns=["Company","Followers","Employees"])

# Gather information of all companies listed on the .txt file
for i in range(number_of_companies):
    link = file_lines[i]
    txt = get_txt_from_screenshot(link)
    company_name = get_company_name(txt)
    number_of_followers = get_number_of_followers(txt)
    number_of_employees = get_number_of_employees(txt)
    entry = pd.DataFrame.from_dict({
        "Company": [company_name],
        "Followers": [number_of_followers],
        "Employees": [number_of_employees]
    })
    dataframe = pd.concat([dataframe, entry], ignore_index=True)

# Define the name of the .xlsx file
date = datetime.today().strftime('%Y-%m-%d')
file_name = "Linkedin Report " + date +".xlsx"

# Generate excel file
dataframe.to_excel(file_name,index=False)