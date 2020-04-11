import bs4
import os
from selenium import webdriver
import datetime
import time
from win32com.client import Dispatch

### Initiate Chrome web driver
### Please include the path to ChromeDriver, download it from https://chromedriver.chromium.org/downloads
driver = webdriver.Chrome('E:\Whole-Foods-Delivery-Slot-master\Whole-Foods-Delivery-Slot-master\chromedriver.exe')

driver.get('https://www.amazon.com/gp/buy/shipoptionselect/handlers/display.html?hasWorkingJavascript=1')

speak = Dispatch("SAPI.SpVoice")

try:
    ### Find sign-in button
    driver.find_element_by_link_text("Sign in").click()

    email = driver.find_element_by_id("ap_email")
    email.clear() 

    ### Input your Amazon account name
    email.send_keys("my account")

    
    driver.find_element_by_id("continue").click()

    ### Input your Amazon password
    password = driver.find_element_by_id("ap_password")
    password.clear() 
    password.send_keys("my password")

    driver.find_element_by_id("signInSubmit").click()
    
    ### Go to shopping cart
    driver.find_element_by_id("nav-cart").click()

    #driver.find_element_by_id("sc-alm-buy-box-ptc-button-QW1hem9uIEZyZXNo").click()

    driver.find_element_by_id("sc-alm-buy-box-ptc-button-VUZHIFdob2xlIEZvb2Rz").click()
except:
    print("Continue")
    speak.Speak("Please do it manually!")

### Continue on page "Before your checkout"
n=1
while n<=3:
    try:
        driver.find_element_by_name("proceedToCheckout").click()
        n=4
    except:
        speak.Speak("Please do step one manually!")
        n+=1
        time.sleep(1)

### Continue on page "Substitution preference"        
n=1
while n<=3:
    try:
        driver.find_element_by_class_name("a-button-inner").click()
        n=4
    except:
        speak.Speak("Please do step two manually!")
        n+=1
        time.sleep(1)

no_open_slots=True

while no_open_slots:
    driver.refresh()
    now=datetime.datetime.now()
    print(now.strftime("%Y-%m-%d %H:%M:%S"),"refreshed")
    html = driver.page_source
    soup = bs4.BeautifulSoup(html)
    time.sleep(4)

    slot_pattern = 'Next available'
    try:
        next_slot_text = soup.find('h4', class_ ='ufss-slotgroup-heading-text a-text-normal').text
        if slot_pattern in next_slot_text:
            print('SLOTS OPEN!')
            speak.Speak("Whole Foods slots for delivery opened!")
            no_open_slots = False
    except AttributeError:
        continue

    try:
        slot_opened_text = "Not available"
        all_dates = soup.findAll("div", {"class": "ufss-date-select-toggle-text-availability"})
        for each_date in all_dates:
            if slot_opened_text not in each_date.text:
                print('SLOTS OPEN!')
                speak.Speak("Whole Foods slots for delivery opened!")
                no_open_slots = False
    except AttributeError:
        continue

    try:
        no_slot_pattern = 'No delivery windows available. New windows are released throughout the day.'
        if no_slot_pattern == soup.find('h4', class_ ='a-alert-heading').text:
            print("NO SLOTS!")
    except AttributeError: 
        print('SLOTS OPEN!')
        speak.Speak("Whole Foods slots for delivery opened!")
        no_open_slots = False

### Find available time slot
n=1
while n<=3:
    try:
        driver.find_element_by_class_name("ufss-slot-price-container").click()
        n=4
    except:
        print("Continue")
        speak.Speak("Please do step one manually!")
        n+=1
        time.sleep(1)

### Click "Continue" to go to checkout
n=1
while n<=3:
    try:
        driver.find_element_by_class_name("a-button-input").click()
        n=4
    except:
        print("Continue")
        speak.Speak("Please do step two manually!")
        n+=1
        time.sleep(1)

### Continue on page "Payment method"
n=1
while n<=3:
    try:
        driver.find_element_by_id("continue-top").click()
        n=4
    except:
        print("Continue")
        speak.Speak("Please do step three manually!")
        n+=1
        time.sleep(1)

### Place your order!
n=1
while n<=3:
    try:
        driver.find_element_by_id("placeYourOrder").click()
        n=4
    except:
        print("Continue")
        speak.Speak("Please do step four manually!")
        n+=1
        time.sleep(1)
    
time.sleep(60)
