import pandas as pd
import datetime
import smtplib

#for whatsapp message
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys


GMAIL_ID = 'Enter the Gmail Id that you wish to send the message from'
GMAIL_PASSWORD = 'Include your Gmail Password'


def send_email(to_email_id, subject, message):
    # print(f"Email to '{to_email_id}' sent with subject '{subject}' with message '{message}' is sent"
         # f" successfully")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASSWORD)

    s.sendmail(GMAIL_ID, to_email_id, f"Subject: {subject} \n\n "
                                      f"This is an auto generated mail from Your_name" 
                                      f"He's just testing his Python application." 
                                      f"Sorry if it disturbed you. {message}")
    s.quit()
    
def send_whatsapp_text(to_number,message,driver):
    #number should be in +91xxxxxxxxxx format
    driver.get('https://web.whatsapp.com/send?phone='+to_number+'&text='+message)
    print("Scan QR Code, And then Enter")
    input_path = '//div[@class="_2S1VP copyable-text selectable-text"][@contenteditable="true"][@data-tab="1"]'
    WebDriverWait(driver,50).until(lambda driver: driver.find_element_by_xpath(input_path)).click()
    driver.find_element_by_xpath(inp_xpath).send_keys(Keys.ENTER)
    

if __name__ == '__main__':

    df = pd.read_excel('Excel_sheet_name.xlsx')
    today = datetime.datetime.now().strftime("%d-%m")
    year_now = datetime.datetime.now().strftime("%Y")
    write_index = []
    
    #change the location to your installed path of chromedriver PN: Download chromedriver according to your chrome version.
    driver = webdriver.Chrome('/Users/nishantkothari/Downloads/chromedriver')
    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        
        if today == bday and year_now not in str(item['Year']):
            send_email(item['Email_Id'], 'Happy Birthday', item['Dialogue'])
            send_whatsapp_text(item['Name'],item['Number'],driver)
            
            write_index.append(index)
    driver.quit()
    
    if len(write_index) > 0:
        for i in write_index:
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + ', ' + str(year_now)
        df.to_excel('Excel_sheet_name.xlsx', index = False)
