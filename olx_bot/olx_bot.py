from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
import time
import telegram


# path to geckodriver
gecko = ''

# link to OLX with filters on
link_olx = "https://www.olx.pl/d/nieruchomosci/mieszkania/wynajem/krakow/?search%5Bdist%5D=5&search%5Bprivate_business%5D=private&search%5Border%5D=created_at:desc&search%5Bfilter_enum_rooms%5D%5B0%5D=two"

# telegram bot info
api_key = ''
user_id = ''
telelist = []



def StartOLX():

    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options, executable_path=gecko)
    driver.get(link_olx)
    listing_link = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[2]/form/div[5]/div/div[2]/div[6]/a")
    listing_time = driver.find_element(By.XPATH,
                                "/html/body/div[1]/div[1]/div[2]/form/div[5]/div/div[2]/div[6]/a/div/div/div[2]/div[2]/p[1]")
    listing_link_2 = listing_link.get_attribute("href")
    listing_time_2 = listing_time.text
    driver.close()

    if listing_link_2 not in telelist:
        bot = telegram.Bot(token=api_key)
        bot.send_message(chat_id=user_id, text=listing_time_2 + "\n" + listing_link_2)
        telelist.append(listing_link_2)


def StartGlobal():
    while True:
        try:
            StartOLX()
        except:
            continue
    time.sleep(11)


StartGlobal()