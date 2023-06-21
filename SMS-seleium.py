from selenium import webdriver
import csv
import time
import sys

from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service



def sendMessagesWithArray(NumberSet,driver):
    count = 0
    driver.implicitly_wait(20)
    driver.get('https://messages.google.com/web/conversations')
    driver.implicitly_wait(20)

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
        contact_input.send_keys(number)
        
        # ---  Click contact selector button ---
        contact_selector = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-contact-selector-button')
        driver.implicitly_wait(5)
        contact_selector.click()

        count = count + 1

    # press next button
    next_btn = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/button')
    driver.implicitly_wait(20)
    next_btn.click()

    # ---  Message input ---
    message = "Enhance your online presence today! \\Enjoy limited-time discounts on our software development packages.\\* FREE web design.\\* FREE domain/hosting for 1 year.\\* Packages start from Rs.2400/month.\\\\We build any web,android,ios,windows applications using the best technologies.\\www.genztech.lk\\\\Reply to this sms and we will get back to you."
    msg_input = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/div[2]/div/div/mws-autosize-textarea/textarea')
    # msg_input.send_keys(message)
    driver.implicitly_wait(20)

    # ---  Send message ---

    send_btn = driver.find_element(By.XPATH,'/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/mws-message-send-button/div/mw-message-send-button')
    driver.implicitly_wait(20)
    send_btn.click()

    print(f"{count} messages sent!")


def get_data_from_csv(filename):
    
  with open(filename, 'r') as csvfile:
    reader = csv.reader(csvfile, delimiter=',')
    data = []
    for row in reader:
      data.append(row[0])
  arrays = []
  for i in range(0, len(data), 28):
    array = data[i:i + 28]
    if len(array) == 28:
      arrays.append(array)
  return arrays

def loadPage():
    print("*************************************************************")
    print("SMS sending program...")
    print("*************************************************************")
    print("\n")

    #  Check here for update
    service = Service('C:\\Users\\Sudeera Perera\\msedgedriver.exe')
    driver = webdriver.Edge(service=service)

    print("Successfully connected.....")

    return driver

def main():
    driver = loadPage()

    filename = 'numbers-entertainment.csv'
    arrays = get_data_from_csv(filename)
    for array in arrays:
        print(array)
        sendMessagesWithArray(array,driver)
        



if __name__ == "__main__":
    main()