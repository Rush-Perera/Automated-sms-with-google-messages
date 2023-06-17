from selenium import webdriver
import openpyxl
import time
import sys

from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service

# check here for update
filex = openpyxl.load_workbook("C:\\Users\\Sudeera Perera\\OneDrive - engug.ruh.ac.lk\\GENZ\Bulk SMS\\SMS-selenium-python\\numbers-test.xlsx")
sh = filex["Sheet1"]

print("*************************************************************")
print("SMS sending program...")
print("*************************************************************")
print("\n")

print("Please enter the limit from excel file :")
st = int(input("1. Start point "))
ed = int(input("2. End point "))

#  Check here for update
service = Service('C:\\Users\\Sudeera Perera\\msedgedriver.exe')
driver = webdriver.Edge(service=service)
driver.get('https://messages.google.com/web/conversations')

time.sleep(8)
print("Successfully scanned.....")

time.sleep(8)
element1 = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/div/mw-fab-link/a')
element1.click()
time.sleep(8)

ele = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-new-conversation-start-group-button/button')
ele.click()
time.sleep(5)
count = 0
while(st<=ed):

    cl = sh.cell(st,1)
    # ---  Contact input ---
    element = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/div/mat-chip-listbox/div/input')
    element.send_keys(cl.value)
    time.sleep(1)
    element2 = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-contact-selector-button/button')
    element2.click()
    # time.sleep(1)
    st+=1
    count+=1
time.sleep(2)
#Add message here
message = "Hello. This is a test message from python."
element3 = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/button')
element3.click()
time.sleep(16)
messend = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/div[2]/div/div/mws-autosize-textarea/textarea')

messend.send_keys(message)
time.sleep(4)

element4 = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/mws-message-send-button/div/mw-message-send-button')
element4.click()
time.sleep(2)
print(f"{count} messages sent already..")

sys.exit() 
