from selenium import webdriver
import openpyxl
import time
import sys

from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service

# check here for update
filex = openpyxl.load_workbook("C:\\Users\\Sudeera Perera\\OneDrive - engug.ruh.ac.lk\\GENZ\Bulk SMS\\SMS-selenium-python\\numbers-test.xlsx")
sheet = filex["Sheet1"]

print("*************************************************************")
print("SMS sending program...")
print("*************************************************************")
print("\n")

print("Please enter the limit from excel file :")
st = int(input("1. Start point "))
ed = int(input("2. End point "))
num_count = ed-st

#  Check here for update
service = Service('C:\\Users\\Sudeera Perera\\msedgedriver.exe')
driver = webdriver.Edge(service=service)
driver.get('https://messages.google.com/web/conversations')
driver.implicitly_wait(20)

print("Successfully scanned.....")

count = 0

while(st<=ed):

    # if the remander of count divided by 29 is not a whole number
    if(is_whole_number(count/5)):
      
        # messend.send_keys(message)
        driver.implicitly_wait(20)

        element4 = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/mws-message-send-button/div/mw-message-send-button')
        driver.implicitly_wait(20)
        element4.click()
        print(f"{count} messages sent!")
        
    


    cl = sheet.cell(st,1)
        # ---  Contact input ---
    
    # time.sleep(1)1
    st+=1
    count+=1
    num_count-=1



def sendMessagesWithArray(NumberSet):
    # Click start chat button
    start_chat_btn = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/div/mw-fab-link/a')
    driver.implicitly_wait(20)
    start_chat_btn.click()

    # Click group chat button
    group_chat_btn = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-new-conversation-start-group-button/button')
    driver.implicitly_wait(20)
    group_chat_btn.click()

    for number in NumberSet:

        # ---  Contact input ---
        contact_input = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/div/mat-chip-listbox/div/input')
        contact_input.send_keys(number.value)
        
        # ---  Click contact selector button ---
        contact_selector = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-contact-selector-button')
        driver.implicitly_wait(5)
        contact_selector.click()

    # press next button
    next_btn = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/button')
    driver.implicitly_wait(20)
    next_btn.click()

    # ---  Message input ---
    message = "Enhance your online presence today! \\Enjoy limited-time discounts on our software development packages.\\* FREE web design.\\* FREE domain/hosting for 1 year.\\* Packages start from Rs.2400/month.\\\\We build any web,android,ios,windows applications using the best technologies.\\www.genztech.lk\\\\Reply to this sms and we will get back to you."
    msg_input = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/div[2]/div/div/mws-autosize-textarea/textarea')
    msg_input.send_keys(message)
    driver.implicitly_wait(20)

    # ---  Send message ---

    send_btn = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/mws-message-send-button/div/mw-message-send-button')
    driver.implicitly_wait(20)
    send_btn.click()

    print(f"{count} messages sent!")